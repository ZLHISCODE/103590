VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmClinicPlanDaysManage 
   BorderStyle     =   0  'None
   Caption         =   "���ﰲ�Ź���"
   ClientHeight    =   7410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   9150
      MaxLength       =   100
      TabIndex        =   13
      Top             =   1140
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.PictureBox picSelectWeek 
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   90
      ScaleHeight     =   345
      ScaleWidth      =   10845
      TabIndex        =   3
      Top             =   450
      Width           =   10845
      Begin VB.CheckBox chkShowAllPlan 
         Caption         =   "��ʾ�������а���"
         Height          =   225
         Left            =   7020
         TabIndex        =   11
         Top             =   75
         Width           =   1755
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "��6��"
         Height          =   195
         Index           =   6
         Left            =   6000
         TabIndex        =   10
         Top             =   90
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "��5��"
         Height          =   195
         Index           =   5
         Left            =   5040
         TabIndex        =   9
         Top             =   90
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "��4��"
         Height          =   195
         Index           =   4
         Left            =   4050
         TabIndex        =   8
         Top             =   90
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "��3��"
         Height          =   195
         Index           =   3
         Left            =   3045
         TabIndex        =   7
         Top             =   90
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "��2��"
         Height          =   195
         Index           =   2
         Left            =   2055
         TabIndex        =   6
         Top             =   90
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "��1��"
         Height          =   195
         Index           =   1
         Left            =   1050
         TabIndex        =   5
         Top             =   90
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "ȫ��"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   90
         Value           =   -1  'True
         Width           =   705
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfRegistPlan 
      Height          =   2445
      Left            =   630
      TabIndex        =   1
      Top             =   1110
      Width           =   3495
      _cx             =   6165
      _cy             =   4313
      Appearance      =   0
      BorderStyle     =   0
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
      FormatString    =   $"frmClinicPlanDaysManage.frx":0000
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
         Picture         =   "frmClinicPlanDaysManage.frx":0075
         ScaleHeight     =   135
         ScaleWidth      =   150
         TabIndex        =   2
         Top             =   90
         Width           =   150
      End
   End
   Begin VB.Label lblPublishInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ˣ�Ƚ����  ����ʱ�䣺2016-01-02 12:32:12"
      Height          =   180
      Left            =   6840
      TabIndex        =   12
      Top             =   150
      Width           =   4050
   End
   Begin VB.Line lineSplit 
      BorderColor     =   &H8000000A&
      X1              =   1020
      X2              =   5010
      Y1              =   960
      Y2              =   960
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   10845
      _Version        =   589884
      _ExtentX        =   19129
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "���ﰲ��>���ﰲ��"
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
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   1035
      Left            =   30
      Top             =   960
      Width           =   525
   End
End
Attribute VB_Name = "frmClinicPlanDaysManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar�ؼ�
Private mlngModule As Long
Private mstrPrivs As String

Private mbytFun As Byte '1-�°��ţ�2-�ܰ���
Private mlng����ID As Long
Private mrsPlanRecords As ADODB.Recordset
Private mintFindType As Integer

Private mlngCopyPlanID As Long, mstrCopyPlanItem As String
Private mintYear As Integer, mintMonth As Integer, mintWeek As Integer
Private mlng�����ܳ���ID As Long, mintElseWeek As Integer '�ܰ��ſ���ʱ���������������¼��һ����������Ϣ
Private mdtStartDate As Date, mdtEndDate As Date

Private mdtToday As Date
Private mstrOldSelRangePlan As String 'ѡ���������򣬸�ʽ"��ʼ��|������|��ʼ��|������"

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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ�������(&D)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddNewSignalSource, "������Դ����(&A)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyPlanItem, "��������(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyUnitRegist, "����ԤԼ�Һſ���(&R)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AllStartNO, "ȫ��������ſ���(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AllStopNO, "ȫ��ȡ����ſ���(&T)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CopyPlan, "���ư���(&C)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PastPlan, "ճ������(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearCurPlan, "�����ǰ����(&C)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearAllPlan, "�����ǰ��Դ����(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearAll, "������к�Դ����(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyToDay, "Ӧ���ڡ����е��ա�(&D)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyToWeekDay, "Ӧ���ڡ��������ڼ���(&W)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PublishPlan, "��������(&G)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnPublishPlan, "ȡ������(&I)")
        
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UpdateUnitRegist, "����ԤԼ�Һſ���(&H)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrintPlan, "��ӡ�����(&P)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PlanSaveAsTemplet, "���Ϊģ��(&A)..."): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextNewPlan, "�����³����(&N)")
    End With

    '�鿴�˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) 'ˢ����ǰ(���ʱע�ⷴ��)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowDoctorStopPlan, "��ʾҽ��ͣ�ﰲ��(&P)", cbrControl.index)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_PlanChangeInfo, "��ѯ�䶯��Ϣ(&C)", cbrControl.index)
        cbrControl.BeginGroup = True
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddMonthPlan, "�³����", cbrControl.index + 1): cbrControl.BeginGroup = True
        cbrControl.ToolTipText = "�ƶ��³����"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddWeekPlan, "�ܳ����", cbrControl.index + 1)
        cbrControl.ToolTipText = "�ƶ��ܳ����"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ�������", cbrControl.index + 1): cbrControl.BeginGroup = True
        
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

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextNewPlan, "�����³����", cbrControl.index + 1): cbrControl.BeginGroup = True
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
        .Add FCONTROL, Asc("D"), conMenu_Edit_Delete
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

Private Function IsInCurrentPlan(ByVal vsfGrid As VSFlexGrid) As Boolean
    '�жϵ�ǰѡ�������Ƿ��ڵ�ǰ������
    Dim strCurItem As String
    On Error GoTo errHandler
    With vsfGrid
        If .Col < gPlanGrid_FixedCols Then Exit Function
        If .Row < .FixedRows Or .Row > .Rows - 1 Then Exit Function
        strCurItem = .Cell(flexcpData, 0, .Col)
        If IsDate(strCurItem) = False Then Exit Function
        If strCurItem < mdtStartDate Or strCurItem > mdtEndDate Then Exit Function
    End With
    IsInCurrentPlan = True
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
    Dim strDoctorName As String
    Dim blnEnabled As Boolean
    
    '˵������ʾ�������а���ʱ��ֻ��˫���鿴�����ܲ����κι���
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    blnEnabled = mlng����ID <> 0
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = vsfRegistPlan.Rows > vsfRegistPlan.FixedRows
    Case conMenu_EditPopup
        If mfrmMain.mFunListActived Then
            Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��;��������;ȡ������")
        Else
            Control.Visible = chkShowAllPlan.Value = vbUnchecked _
                And HavePrivs(mstrPrivs, "���ﰲ��;��������;��ʱ���ﰲ��;ͣ��;����;�Ӻ�;����;������������;����ԤԼ�Һ�;ģ�����")
        End If
        Control.Enabled = Control.Visible
    Case conMenu_Edit_AddMonthPlan, conMenu_Edit_AddWeekPlan '�ƶ��³����,�ƶ��ܳ����
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And mfrmMain.mFunListActived
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Delete 'ɾ�������
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And mfrmMain.mFunListActived
        Control.Enabled = Control.Visible And blnEnabled And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
    Case conMenu_Edit_PublishPlan, conMenu_Edit_UnPublishPlan '��������,ȡ������
        Control.Visible = mfrmMain.mFunListActived And HavePrivs(mstrPrivs, Decode(Control.ID, _
            conMenu_Edit_PublishPlan, "��������", conMenu_Edit_UnPublishPlan, "ȡ������")) And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then
            If Control.ID = conMenu_Edit_PublishPlan Then
                blnEnabled = Val(lblPublishInfo.Tag) = 0
            Else
                blnEnabled = Val(lblPublishInfo.Tag) = 1
            End If
        End If
        Control.Enabled = Control.Visible And blnEnabled
    
    Case conMenu_Edit_ModifyPlanItem '����������
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsCan�ٴ��Ű�(vsfRegistPlan)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ModifyUnitRegist '����ԤԼ�Һſ���
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsCan�ٴ��Ű�(vsfRegistPlan)
        If blnEnabled Then blnEnabled = Is��ֹԤԼ(vsfRegistPlan) = False
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_AllStartNO, conMenu_Edit_AllStopNO 'ȫ��������ſ���,ȫ��ȡ����ſ���
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_CopyPlan '���ư���
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_PastPlan 'ճ������
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsCan�ٴ��Ű�(vsfRegistPlan)
        Control.Enabled = Control.Visible And blnEnabled And mlngCopyPlanID <> 0
    Case conMenu_Edit_ClearCurPlan, conMenu_Edit_ClearAllPlan '�����ǰ����,�����ǰ��Դ���а���
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
        If Control.ID = conMenu_Edit_ClearCurPlan And blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsCan�ٴ��Ű�(vsfRegistPlan)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ClearAll '������к�Դ����
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ApplyToDay, conMenu_Edit_ApplyToWeekDay 'Ӧ���ڡ����е��ա�,Ӧ���ڡ��������ڼ���
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
        If Control.ID = conMenu_Edit_ApplyToWeekDay And Control.Visible Then Control.Visible = mbytFun = 1
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsCan�ٴ��Ű�(vsfRegistPlan)
        Control.Enabled = Control.Visible And blnEnabled

    '�ѷ������ŵ���
    Case conMenu_Edit_AddNewSignalSource '������Դ����
        Control.Visible = HavePrivs(mstrPrivs, "��������") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 1 And chkShowAllPlan.Value = vbUnchecked
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_LockResource, conMenu_Edit_UnLockResource '����,����
        Control.Visible = HavePrivs(mstrPrivs, "��������") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 1 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsValid(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then
            If Control.ID = conMenu_Edit_LockResource Then
                blnEnabled = (PlanIsSelOne(vsfRegistPlan) = False Or PlanIsLocked(vsfRegistPlan) = False)
            Else
                blnEnabled = (PlanIsSelOne(vsfRegistPlan) = False Or PlanIsLocked(vsfRegistPlan))
            End If
        End If
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_AddTempPlan, conMenu_Edit_UpdatePlan '��ʱ����,����������İ���
        Control.Visible = HavePrivs(mstrPrivs, "��������") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 1 And chkShowAllPlan.Value = vbUnchecked
        If Control.ID = conMenu_Edit_UpdatePlan And blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsSelOne(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsValid(vsfRegistPlan)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_StopOutCall, conMenu_Edit_UnStopOutCall, conMenu_Edit_OpenStopPlan 'ͣ��,ȡ��ͣ��,����ͣ�ﰲ��
        Control.Visible = HavePrivs(mstrPrivs, "ͣ��") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 1 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsValid(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsSelOne(vsfRegistPlan)
        If blnEnabled Then
            If Control.ID = conMenu_Edit_StopOutCall Then
                blnEnabled = (PlanIsStopVisit(vsfRegistPlan) = False)
            Else
                blnEnabled = PlanIsStopVisit(vsfRegistPlan)
            End If
        End If
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ModifyDoctor, conMenu_Edit_UnModifyDoctor '����,ȡ������
        Control.Visible = HavePrivs(mstrPrivs, "����") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 1 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsValid(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsSelOne(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsStopVisit(vsfRegistPlan) = False
        If blnEnabled Then
            If Control.ID = conMenu_Edit_ModifyDoctor Then
                blnEnabled = (PlanIsReplaceDoctor(vsfRegistPlan) = False)
            Else
                blnEnabled = PlanIsReplaceDoctor(vsfRegistPlan)
            End If
        End If
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_AddNumberLimit, conMenu_Edit_ReduceNumberLimit, _
        conMenu_Edit_ModifyDoctorOffice, conMenu_Edit_UpdateUnitRegist '�Ӻ�,����,������������,����ԤԼ�Һſ���
        Control.Visible = HavePrivs(mstrPrivs, Decode(Control.ID, _
            conMenu_Edit_AddNumberLimit, "�Ӻ�", conMenu_Edit_ReduceNumberLimit, "����", _
            conMenu_Edit_ModifyDoctorOffice, "������������", conMenu_Edit_UpdateUnitRegist, "����ԤԼ�Һ�")) _
            And mfrmMain.mFunListActived = False And Val(lblPublishInfo.Tag) = 1 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsValid(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsSelOne(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsStopVisit(vsfRegistPlan) = False
        If Control.ID = conMenu_Edit_UpdateUnitRegist And blnEnabled Then blnEnabled = Is��ֹԤԼ(vsfRegistPlan) = False
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_NextNewPlan, conMenu_Edit_PlanSaveAsTemplet '�����°���,���Ϊģ��(&A)...
        Control.Visible = chkShowAllPlan.Value = vbUnchecked And HavePrivs(mstrPrivs, Decode(Control.ID, _
            conMenu_Edit_NextNewPlan, "���ﰲ��", conMenu_Edit_PlanSaveAsTemplet, "ģ�����"))
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_PrintPlan '    ��ӡ�����
        Control.Visible = HavePrivs(mstrPrivs, Decode(mbytFun, 1, "�³����", "�ܳ����")) And chkShowAllPlan.Value = vbUnchecked
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_View_FindType '���ҷ�ʽ
        Control.Caption = "��" & Decode(mintFindType, 0, "����", 1, "����", 2, "ҽ��", "����") & "���ˡ�"
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 9 '���ҷ�ʽ
        Control.Checked = Val(Right(Control.ID, 2)) - 1 = mintFindType
    Case conMenu_View_ShowDoctorStopPlan '��ʾҽ��ͣ�ﰲ��
        Control.Visible = mfrmMain.mFunListActived = False
        blnEnabled = False
        If vsfRegistPlan.Row >= vsfRegistPlan.FixedRows Then
            blnEnabled = Trim(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, COL_ҽ��)) <> ""
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
    Dim lng����ID As Long
    Dim lng��¼ID As Long, lng��ԴId As Long, str���� As String, strItem As String
    Dim obj�����¼ As �����¼, obj�����Դ As �����Դ
    Dim blnFixedRule As Boolean
    Dim strIDs As String, lngRowStart As Long, lngRowEnd As Long, i As Integer
    Dim lngCurCol As Long
    Dim strKey As String
    Dim strDoctorName As String
    Dim str��¼IDs As String
    
    Err = 0: On Error GoTo errHandler
    lng����ID = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_����ID))
    strItem = vsfRegistPlan.Cell(flexcpData, 0, vsfRegistPlan.Col)
    strDoctorName = Trim(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, COL_ҽ��))
    
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_Delete 'ɾ�������
        If DeletePlan(mlng����ID, mbytFun, mintWeek, mlng�����ܳ���ID, mintElseWeek) Then Call mfrmMain.NodeChanged("")
    Case conMenu_Edit_ModifyPlanItem '��������
        lng��ԴId = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ԴID))
        If lng��ԴId <> 0 Or lng����ID <> 0 Then
            Set frmEdit = New frmClinicPlanEdit
            If frmEdit.ShowMe(Me, mbytFun, Fun_Update, mlng����ID, lng��ԴId, lng����ID, strItem) Then
                Call RefreshOneData
                mlngCopyPlanID = 0: mstrCopyPlanItem = ""
            End If
        End If
    Case conMenu_Edit_ModifyUnitRegist '����ԤԼ�Һſ���
        lng��ԴId = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ԴID))
        If lng��ԴId <> 0 Or lng����ID <> 0 Then
            Set frmEdit = New frmClinicPlanEdit
            Call frmEdit.ShowMe(Me, IIf(mbytFun = 1, 1, 2), Fun_UpdateUnit, mlng����ID, lng��ԴId, lng����ID, strItem)
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
        If PastPlan(mlng����ID, mlngCopyPlanID, mstrCopyPlanItem) Then Call RefreshOneData
    Case conMenu_Edit_ClearCurPlan '�����ǰ����
        If strItem = "" Then Exit Sub
        If IsDate(strItem) = False Then Exit Sub
        str���� = vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_����)
        If MsgBox("��ȷ��Ҫ�������Ϊ��" & str���� & "����" & Format(strItem, "mm��dd��") & "���İ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        If ZlClearPlan(lng����ID, strItem, True) Then
            Call RefreshOneData
            If mlngCopyPlanID = lng����ID And mstrCopyPlanItem = strItem Then
                mlngCopyPlanID = 0: mstrCopyPlanItem = ""
            End If
        End If
    Case conMenu_Edit_ClearAllPlan '�����ǰ��Դ����
        str���� = vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_����)
        If MsgBox("��ȷ��Ҫ�������Ϊ��" & str���� & "�������а�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        lng��ԴId = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ԴID))
        If ZlClearPlanBatch(mlng����ID, lng��ԴId) Then
            Call RefreshOneData
            If mlngCopyPlanID = lng����ID Then
                mlngCopyPlanID = 0: mstrCopyPlanItem = ""
            End If
        End If
    Case conMenu_Edit_ClearAll '������к�Դ����
        If MsgBox("��ȷ��Ҫ������к�Դ�İ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        If ZlClearPlanBatch(mlng����ID, 0, IIf(HavePrivs(mstrPrivs, "���п���"), 0, UserInfo.ID)) Then
            Call RefreshData(mbytFun, mlng����ID)
            mlngCopyPlanID = 0: mstrCopyPlanItem = ""
        End If
    Case conMenu_Edit_ApplyToDay 'Ӧ���ڡ����е��ա�
        If ApplyToDay(lng����ID, strItem) Then Call RefreshOneData
    Case conMenu_Edit_ApplyToWeekDay 'Ӧ���ڡ��������ڼ���
        If ApplyToWeekDay(lng����ID, strItem) Then Call RefreshOneData
    Case conMenu_Edit_PublishPlan '��������
        If PublishPlan(mlng����ID, True, mlng�����ܳ���ID) Then
            Call PrintPlan(mlng����ID, mbytFun)
            'ˢ������
            If mbytFun = 1 Then
                strKey = "K2_" & mintYear & "_" & mintMonth 'XX�³����ڵ�: K2_���_�·�
            Else
                strKey = "K3_" & mintYear & "_" & mintMonth & "_" & mintWeek 'XX�ܳ����ڵ㣺K3_���_�·�_����
            End If
            Call mfrmMain.NodeChanged(strKey)
        End If
    Case conMenu_Edit_UnPublishPlan 'ȡ������
        If PublishPlan(mlng����ID, False, mlng�����ܳ���ID) Then
            'ˢ������
            If mbytFun = 1 Then
                strKey = "K2_" & mintYear & "_" & mintMonth 'XX�³����ڵ�: K2_���_�·�
            Else
                strKey = "K3_" & mintYear & "_" & mintMonth & "_" & mintWeek 'XX�ܳ����ڵ㣺K3_���_�·�_����
            End If
            Call mfrmMain.NodeChanged(strKey)
        End If
    
    '�ѷ������ŵ���
    Case conMenu_Edit_AddNewSignalSource '������Դ����
        Set frmEdit = New frmClinicPlanEdit
        If frmEdit.ShowMe(Me, mbytFun, Fun_AddSignalSourcePlan, mlng����ID, , , strItem) Then
            Call RefreshData(mbytFun, mlng����ID)
        End If
    Case conMenu_Edit_LockResource '����
        Call LockPlan(False)
    Case conMenu_Edit_UnLockResource '����
        Call LockPlan(True)
    Case conMenu_Edit_AddTempPlan '��ʱ����
        Set frmEdit = New frmClinicPlanEdit
        lng��ԴId = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ԴID))
        If frmEdit.ShowMe(Me, mbytFun, Fun_TempPlanRecord, mlng����ID, lng��ԴId, lng����ID, strItem) Then
            Call RefreshOneData
        End If
    Case conMenu_Edit_UpdatePlan '����������İ���
        lng��ԴId = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ԴID))
        If lng��ԴId = 0 Or lng����ID = 0 Then Exit Sub
        
        Call LockPlanByDay(False, str��¼IDs)
        Set frmEdit = New frmClinicPlanEdit
        If frmEdit.ShowMe(Me, mbytFun, Fun_UpdatePlan, mlng����ID, lng��ԴId, lng����ID, strItem) Then
            Call LockPlanByDay(True, str��¼IDs)
            Call RefreshOneData
        End If
        Call LockPlanByDay(True, str��¼IDs)
    Case conMenu_Edit_StopOutCall 'ͣ��
        lng��¼ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
        If lng��¼ID = 0 Then Exit Sub
        Set frmStopVisitAndModifyDoctor = New frmClinicPlanStopVisitAndModifyDoctor
        'bytFun ���ܣ�1-ͣ��,2-ȡ��ͣ��,3-����,4-ȡ������
        If frmStopVisitAndModifyDoctor.ShowMe(Me, mlngModule, 1, lng��¼ID) Then
            Call RefreshOneData
        End If
    Case conMenu_Edit_UnStopOutCall 'ȡ��ͣ��
        lng��¼ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
        If lng��¼ID = 0 Then Exit Sub
        Set frmStopVisitAndModifyDoctor = New frmClinicPlanStopVisitAndModifyDoctor
        'bytFun ���ܣ�1-ͣ��,2-ȡ��ͣ��,3-����,4-ȡ������
        If frmStopVisitAndModifyDoctor.ShowMe(Me, mlngModule, 2, lng��¼ID) Then
            Call RefreshOneData
        End If
    Case conMenu_Edit_OpenStopPlan '����ͣ�ﰲ��
        lng��¼ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
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
        lng��¼ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
        If lng��¼ID = 0 Then Exit Sub
        Set frmStopVisitAndModifyDoctor = New frmClinicPlanStopVisitAndModifyDoctor
        'bytFun ���ܣ�1-ͣ��,2-ȡ��ͣ��,3-����,4-ȡ������
        If frmStopVisitAndModifyDoctor.ShowMe(Me, mlngModule, 3, lng��¼ID) Then
            Call RefreshOneData
        End If
    Case conMenu_Edit_UnModifyDoctor 'ȡ������
        lng��¼ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
        If lng��¼ID = 0 Then Exit Sub
        Set frmStopVisitAndModifyDoctor = New frmClinicPlanStopVisitAndModifyDoctor
        'bytFun ���ܣ�1-ͣ��,2-ȡ��ͣ��,3-����,4-ȡ������
        If frmStopVisitAndModifyDoctor.ShowMe(Me, mlngModule, 4, lng��¼ID) Then
            Call RefreshOneData
        End If
    Case conMenu_Edit_AddNumberLimit '�Ӻ�
        lng��ԴId = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ԴID))
        lng��¼ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
        If Get�����¼(lng��ԴId, lng��¼ID, True, obj�����Դ, obj�����¼) Then
            Set frmNumberLimitModify = New frmClinicPlanNumberLimitModify
            If frmNumberLimitModify.ShowMe(Me, 1, obj�����Դ, obj�����¼) Then
                Call RefreshOneData
            End If
        End If
    Case conMenu_Edit_ReduceNumberLimit '����
        lng��ԴId = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ԴID))
        lng��¼ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
        If Get�����¼(lng��ԴId, lng��¼ID, True, obj�����Դ, obj�����¼) Then
            Set frmNumberLimitModify = New frmClinicPlanNumberLimitModify
            If frmNumberLimitModify.ShowMe(Me, 2, obj�����Դ, obj�����¼) Then
               Call RefreshOneData
            End If
        End If
    Case conMenu_Edit_ModifyDoctorOffice '������������
        lng��ԴId = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ԴID))
        lng��¼ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
        If Get�����¼(lng��ԴId, lng��¼ID, True, obj�����Դ, obj�����¼) Then
            Set frmOfficeAndUnitRegModify = New frmClinicPlanOfficeAndUnitRegModify
            Call frmOfficeAndUnitRegModify.ShowMe(Me, 1, obj�����Դ, obj�����¼, True)
        End If
    Case conMenu_Edit_UpdateUnitRegist '����ԤԼ�Һſ���
        lng��ԴId = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ԴID))
        lng��¼ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
        If Get�����¼(lng��ԴId, lng��¼ID, True, obj�����Դ, obj�����¼) Then
            Set frmOfficeAndUnitRegModify = New frmClinicPlanOfficeAndUnitRegModify
            Call frmOfficeAndUnitRegModify.ShowMe(Me, 2, obj�����Դ, obj�����¼, True)
        End If
    Case conMenu_Edit_PlanSaveAsTemplet '���Ϊģ��(&A)...
        Call SaveAsTemplet(mlng����ID, mbytFun = 1)
    Case conMenu_Edit_NextNewPlan '�����°���
        Call NextNewPlanByPlan(mlng����ID, mbytFun = 1)
    Case conMenu_Edit_PrintPlan '    ��ӡ�����
        Call PrintPlan(mlng����ID, mbytFun, 1)
    Case conMenu_View_Refresh
        Call RefreshData(mbytFun, mlng����ID)
    Case conMenu_View_PlanChangeInfo '��ѯ��Ϣ
        Dim frmPlanChangeHistory As New frmClinicPlanChangeHistory
        frmPlanChangeHistory.ShowMe Me, mlngModule
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 3 '���ҷ�ʽ
        mintFindType = Val(Right(Control.ID, 2)) - 1
        mcbsMain.RecalcLayout
        txtFind(1).Text = ""
        If txtFind(1).Visible And txtFind(1).Enabled Then txtFind(1).SetFocus
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

Private Sub PrintPlan(ByVal lng����ID As Long, ByVal bytFun As Byte, Optional ByVal bytMode As Byte)
    '��ӡ�����
    '��Σ�
    '   mbytFun 1-�°��ţ�2-�ܰ���
    '   bytMode 0-�������ӡ,1-�˵�ѡ���ӡ
    Dim str����ID As String
    
    Err = 0: On Error GoTo errHandler
    If bytMode = 1 Then '��ֹ�����
        If MsgBox("Ҫ��ӡ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    If bytFun = 1 Then
        If gVisitPlan_ModulePara.byt������ӡ��ʽ = 1 Or bytMode = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_2", Me, "����ID=" & mlng����ID, 2)
        ElseIf gVisitPlan_ModulePara.byt������ӡ��ʽ = 2 Then
            If MsgBox("Ҫ��ӡ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_2", Me, "����ID=" & mlng����ID, 2)
            End If
        End If
    Else
        str����ID = mlng����ID
        If mlng�����ܳ���ID <> 0 Then str����ID = str����ID & "," & mlng�����ܳ���ID
        If gVisitPlan_ModulePara.byt������ӡ��ʽ = 1 Or bytMode = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_3", Me, "����ID=" & str����ID, 2)
        ElseIf gVisitPlan_ModulePara.byt������ӡ��ʽ = 2 Then
            If MsgBox("Ҫ��ӡ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_3", Me, "����ID=" & str����ID, 2)
            End If
        End If
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ExecuteFilter()
    '��������
    Dim strKey As String
    
    Err = 0: On Error GoTo errHandler
    Call zlControl.TxtSelAll(txtFind(1))
    
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
    If mintFindType = 8 Then mintFindType = 0 '���
    Call LoadPlanDataByRecordset(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, mrsPlanRecords, mbytFun, , , _
        Val(lblPublishInfo.Tag) = 1, Format(mdtStartDate, "yyyy-mm-dd"), Format(mdtEndDate, "yyyy-mm-dd"))
    Exit Sub
errHandler:
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
            Call RefreshOneData(i, i = lngRowStart)
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

Private Function DeletePlan(ByVal lng����ID As Long, ByVal byt�Ű෽ʽ As Byte, ByVal intWeek As Integer, _
    ByVal lng�����ܳ���ID As Long, ByVal intElseWeek As Integer) As Boolean
    '���ܣ�ɾ�������
    '��Σ�
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String
    Dim strElsePlanName As String
    Dim strToolTip  As String
    Dim lngArray����ID(1) As Long, i As Integer
    
    Err = 0: On Error GoTo errHandler
    strSQL = "Select 1 From �ٴ������ Where ID = [1] And ����ʱ�� Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If rsTemp.EOF Then
        MsgBox "��ǰ������ѱ�����ɾ�����ѷ���������ɾ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���ܿ��µ��ܳ�����ͬ������
    If byt�Ű෽ʽ = 2 And lng�����ܳ���ID <> 0 Then
        strSQL = "Select �������,����ʱ�� From �ٴ������ Where ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�����ܳ���ID)
        If rsTemp.EOF Then
            lng�����ܳ���ID = 0
        Else
            If Nvl(rsTemp!����ʱ��) <> "" Then
                MsgBox "��ǰ��������������ڵ���һ��������ѷ���������ɾ����", vbInformation, gstrSysName
                Exit Function
            End If
            strElsePlanName = Nvl(rsTemp!�������)
        End If
    End If
    
    strSQL = "Select ID" & vbNewLine & _
            " From (Select a.Id" & vbNewLine & _
            "       From �ٴ������ A, �ٴ����ﰲ�� B" & vbNewLine & _
            "       Where a.�Ű෽ʽ = [1] And a.Id = b.����id(+) And Nvl(a.վ��,'-') = Nvl([2],'-')" & vbNewLine & _
            "       Order By a.��� Desc, a.�·� Desc, a.���� Desc)" & vbNewLine & _
            " Where Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, byt�Ű෽ʽ, gstrNodeNo)
    If rsTemp.EOF Then
        MsgBox "��ǰ��������ڣ������ѱ�����ɾ������ˢ�²鿴��", vbInformation, gstrSysName
        Exit Function
    Else
        If Val(Nvl(rsTemp!ID)) <> lng����ID _
            And (byt�Ű෽ʽ = 1 Or byt�Ű෽ʽ = 2 And Val(Nvl(rsTemp!ID)) <> lng�����ܳ���ID) Then
            MsgBox "ɾ��ʧ�ܣ���ֻ�ܴ����һ��δ�����ĳ����ʼɾ����", vbInformation, gstrSysName: Exit Function
            Exit Function
        End If
    End If
    
    If zlStr.IsHavePrivs(mstrPrivs, "���п���") = False Then
        '���������к�������������Ա���Ű�Ͳ���ɾ���ó����
        strSQL = "Select 1" & vbNewLine & _
                " From �ٴ����ﰲ�� A, �ٴ������Դ B" & vbNewLine & _
                " Where a.��Դid = b.Id And a.����id In ([1],[2])" & vbNewLine & _
                "       And Not (Nvl(b.�Ƿ��ٴ��Ű�, 0) = 1 And Exists (Select 1 From ������Ա Where ����id = b.����id And ��Աid = [3]))"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng�����ܳ���ID, UserInfo.ID)
        If Not rsTemp.EOF Then
            MsgBox "��ǰ������к���������Ա�Ѿ��ƶ��İ��ţ�����ɾ����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    strToolTip = "��" & Mid(sccTitle.Caption, InStr(sccTitle.Caption, ">") + 1) & "��"
    If byt�Ű෽ʽ = 2 And lng�����ܳ���ID <> 0 Then
        If intWeek > intElseWeek Then
            strToolTip = strToolTip & "�͡�" & strElsePlanName & "��"
            lngArray����ID(0) = lng�����ܳ���ID
            lngArray����ID(1) = lng����ID
        Else
            strToolTip = "��" & strElsePlanName & "����" & strToolTip
            lngArray����ID(0) = lng����ID
            lngArray����ID(1) = lng�����ܳ���ID
        End If
    End If
    If MsgBox("��ȷ��Ҫɾ��" & strToolTip & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Function
    End If
    
    'ɾ�������
    If byt�Ű෽ʽ = 2 And lng�����ܳ���ID <> 0 Then
        On Error GoTo TransErrHandler
        gcnOracle.BeginTrans
            For i = 0 To UBound(lngArray����ID)
                If lngArray����ID(i) <> 0 Then
                    'Zl_�ٴ������_Delete
                    strSQL = "Zl_�ٴ������_Delete("
                    '  Id_In       �ٴ������.Id%Type
                    strSQL = strSQL & "" & lngArray����ID(i) & ","
                    '  ��Աid_In ��Ա��.Id%Type := Null
                    strSQL = strSQL & "" & IIf(zlStr.IsHavePrivs(mstrPrivs, "���п���"), 0, UserInfo.ID) & ","
                    '  վ��_In   ���ű�.վ��%Type
                    strSQL = strSQL & "'" & gstrNodeNo & "')"
                    zlDatabase.ExecuteProcedure strSQL, "ɾ�������"
                End If
            Next
        gcnOracle.CommitTrans
        DeletePlan = True
    Else
        DeletePlan = ZlDeletePlan(lng����ID, IIf(zlStr.IsHavePrivs(mstrPrivs, "���п���"), 0, UserInfo.ID))
    End If
    Exit Function
TransErrHandler:
    gcnOracle.RollbackTrans
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub RefreshData(ByVal bytFun As Byte, ByVal lng����ID As Long, Optional ByVal blnClear As Boolean, _
    Optional ByVal intYear As Integer, Optional ByVal intMonth As Integer, Optional ByVal strTitle As String)
    '���ܣ�ˢ�°�����������
    '��Σ�
    '   bytFun - 1-�°��ţ�2-�ܰ���
    '   lng����ID - ����ID
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim i As Integer, varDateRange As Variant
    Dim dtStartDate As Date, dtEndDate As Date
    Dim intWeek As Integer
    
    Err = 0: On Error GoTo errHandler
    
    mdtToday = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd"))
    If blnClear Then
        mbytFun = bytFun: mlng����ID = lng����ID
        sccTitle.Caption = "���ﰲ��>" & IIf(strTitle = "", "���ﰲ��", strTitle) & IIf(lng����ID = 0, "(�޳����)", "")
        
        chkShowAllPlan.Value = vbUnchecked
        chkShowAllPlan.Visible = bytFun = 1
        picSelectWeek.Visible = bytFun = 1
        For i = optWeek.LBound To optWeek.UBound
            optWeek(i).Visible = True
        Next
        optWeek(0).Value = True
        
        lblPublishInfo.Visible = chkShowAllPlan.Value = vbUnchecked
        lblPublishInfo.Tag = ""
        Set mrsPlanRecords = Nothing
        mlngCopyPlanID = 0: mstrCopyPlanItem = ""
        mintYear = 0: mintMonth = 0: mintWeek = 0
        mlng�����ܳ���ID = 0: mintElseWeek = 0
        
        '�ı�˵�����
        Call ZlUpdatePlanMenu(Me, mcbsMain, bytFun, IIf(HavePrivs(mstrPrivs, "���п���"), 0, UserInfo.ID))
        
        strSQL = "Select b.�������, b.���, b.�·�, b.����, b.������, b.����ʱ��" & vbNewLine & _
                " From �ٴ������ B" & vbNewLine & _
                " Where b.Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������Ϣ", lng����ID)
        If Not rsTemp.EOF Then
            sccTitle.Caption = "���ﰲ��>" & Nvl(rsTemp!�������)
            lblPublishInfo.Tag = IIf(Nvl(rsTemp!����ʱ��) = "", "", "1") '����Ƿ񷢲�
            lblPublishInfo.Caption = "�����ˣ�" & IIf(Nvl(rsTemp!������) = "", "      ", Nvl(rsTemp!������)) & _
                "  ����ʱ�䣺" & IIf(Nvl(rsTemp!����ʱ��) = "", "                   ", Format(Nvl(rsTemp!����ʱ��), "yyyy-mm-dd hh:mm:ss"))
            mintYear = Val(Nvl(rsTemp!���))
            mintMonth = Val(Nvl(rsTemp!�·�))
            mintWeek = Val(Nvl(rsTemp!����))
        End If
        If mintYear = 0 Then mintYear = intYear
        If mintMonth = 0 Then mintMonth = intMonth
        
        If mintYear = 0 Then mintYear = Year(mdtToday)
        If mintMonth = 0 Then mintMonth = Month(mdtToday)
        If mintWeek = 0 Then mintWeek = GetDateWeek(mdtToday)
        
        '����ȷ��
        For i = GetWeekCount(mintYear, mintMonth) + 1 To optWeek.UBound
            optWeek(i).Visible = False
        Next
        
        '��ʾ���ڷ�Χȷ��
        varDateRange = GetDateRange(mintYear, mintMonth, IIf(bytFun = 2, mintWeek, 0))
        mdtStartDate = varDateRange(0): mdtEndDate = varDateRange(1)
        
        dtStartDate = mdtStartDate: dtEndDate = mdtEndDate
        If bytFun = 2 Then
            If IsDoubleMonthWeekPlan(intYear, intMonth, intWeek, dtStartDate, dtEndDate) Then
                mintElseWeek = intWeek
                mlng�����ܳ���ID = Get�ܳ����ID(intYear, intMonth, intWeek)
            End If
        End If
        
        Call InitPlanGrid(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, dtStartDate, dtEndDate, Val(lblPublishInfo.Tag) = 1)
        Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistPlan, Me.Name, "����")
        Call ShowHolidayToPlan(vsfRegistPlan, dtStartDate, dtEndDate)
        Call Form_Resize
    End If
    
    Screen.MousePointer = vbHourglass
    If lng����ID = 0 And chkShowAllPlan.Value = vbUnchecked Then
        Set mrsPlanRecords = Nothing
    Else
        Set mrsPlanRecords = GetPlanRecords(bytFun = 1, lng����ID, Val(lblPublishInfo.Tag) = 1, _
            chkShowAllPlan.Value = vbChecked, mintYear, mintMonth, mdtStartDate, mdtEndDate, mlng�����ܳ���ID)
    End If
    '��������
    Call ExecuteFilter
'    Call ShowStopVisitPlan(vsfRegistPlan, mdtStartDate, mdtEndDate)
    '��λ����ǰ����
    If bytFun = 1 And (mdtToday >= mdtStartDate And mdtStartDate <= mdtEndDate) Then
        For i = gPlanGrid_FixedCols To vsfRegistPlan.Cols - 1 Step 3
            If CDate(vsfRegistPlan.Cell(flexcpData, 0, i)) = mdtToday Then
                vsfRegistPlan.LeftCol = i
                vsfRegistPlan.Col = i
                If vsfRegistPlan.Rows > 2 Then vsfRegistPlan.Row = 3
                Exit For
            End If
        Next
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshOneData(Optional ByVal lngCurRow As Long = -1, _
    Optional ByVal blnReLoadData As Boolean = True)
    'ˢ��ָ���к�Դ����
    Dim lng��ԴId As Long, str�շ���Ŀ As String
    
    Err = 0: On Error GoTo errHandle
    '1.��¼ԭ���ݣ�����ȡ������
    With vsfRegistPlan
        lng��ԴId = Val(.TextMatrix(IIf(lngCurRow = -1, .Row, lngCurRow), COL_��ԴID))
        str�շ���Ŀ = .TextMatrix(IIf(lngCurRow = -1, .Row, lngCurRow), COL_��Ŀ)
    End With
    
    If blnReLoadData Then
        '���±��ؼ�¼��
        Set mrsPlanRecords = GetPlanRecords(mbytFun = 1, mlng����ID, Val(lblPublishInfo.Tag) = 1, _
            chkShowAllPlan.Value = vbChecked, mintYear, mintMonth, mdtStartDate, mdtEndDate, mlng�����ܳ���ID)
    End If
    
    '2.���½���
    mrsPlanRecords.Filter = "��ԴID=" & lng��ԴId & " And �շ���Ŀ='" & str�շ���Ŀ & "'"
    Call RefreshOnePlanData(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, mrsPlanRecords, lngCurRow, _
        Val(lblPublishInfo.Tag) = 1, mbytFun, Format(mdtStartDate, "yyyy-mm-dd"), Format(mdtEndDate, "yyyy-mm-dd"))
'    Call ShowStopVisitPlan(vsfRegistPlan, mdtStartDate, mdtEndDate, lng��ԴId)
    mrsPlanRecords.Filter = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function AddNewPlan(Optional blnMonth As Boolean) As String
    '���ܣ����������
    '��Σ�
    '   blnMonth �Ƿ����Ű�
    Dim strSQL As String, rsTemp As ADODB.Recordset, lng����ID As Long
    Dim intYear As Integer, intMonth As Integer, intWeek As Integer
    Dim dtStart As Date, dtEnd As Date
    Dim strName As String, strKey As String, blnDeletePlan As Boolean
    Dim cllPlan As Collection, i As Integer
    Dim dtCurrent As Date
    
    Err = 0: On Error GoTo errHandler
    Set cllPlan = GetNewPlanInfo(Me, mstrPrivs, blnMonth, strKey, blnDeletePlan)
    If cllPlan Is Nothing Then Exit Function
    If cllPlan.Count = 0 Then Exit Function
    
    dtCurrent = zlDatabase.Currentdate
    On Error GoTo TransErrHandler
    If cllPlan.Count > 1 Then gcnOracle.BeginTrans
    For i = 1 To cllPlan.Count
        'Array(���,�·�,����,��ʼ����,��������)
        intYear = cllPlan(i)(0)
        intMonth = cllPlan(i)(1)
        intWeek = cllPlan(i)(2)
        dtStart = cllPlan(i)(3)
        dtEnd = cllPlan(i)(4)
        
        'ȷ�������ID
        strSQL = "Select ID From �ٴ������ Where �Ű෽ʽ = [1] And ��� = [2] And �·� = [3]" & _
            IIf(blnMonth, "", " And ���� = [4]") & " And Nvl(վ��,'-') = Nvl([5],'-')"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(blnMonth, 1, 2), intYear, intMonth, intWeek, gstrNodeNo)
        If rsTemp.EOF Then
            lng����ID = zlDatabase.GetNextId("�ٴ������")
        Else
            lng����ID = Val(Nvl(rsTemp!ID))
        End If
        
        '�������
        strName = intYear & "��" & intMonth & "��"
        If Not blnMonth Then strName = strName & "��" & intWeek & "��"
        strName = strName & "�����"
        
        'Zl_�ٴ������_Add(
        strSQL = "Zl_�ٴ������_Add("
        '  ��������_In Number,--1-ģ�壬2-�̶�����, 3-�°��ţ�4-�ܰ���
        strSQL = strSQL & "" & IIf(blnMonth, 3, 4) & ","
        '  ����id_In   �ٴ������.Id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  �������_In �ٴ������.�������%Type,
        strSQL = strSQL & "'" & strName & "',"
        '  վ��_In     ���ű�.վ��%Type,
        strSQL = strSQL & "'" & gstrNodeNo & "',"
        '  ȫԺ��Դ����վ��_In ���ű�.վ��%Type,
        strSQL = strSQL & "'" & gVisitPlan_ModulePara.str��Դά��վ�� & "',"
        '  ����Ա_In   �ٴ����ﰲ��.����Ա����%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  ����ʱ��_In �ٴ����ﰲ��.�Ǽ�ʱ��%Type := Null
        strSQL = strSQL & "" & ZDate(dtCurrent) & ","
        '  ��ʼʱ��_In �ٴ����ﰲ��.��ʼʱ��%Type := Null,
        strSQL = strSQL & "" & ZDate(dtStart) & ","
        '  ��ֹʱ��_In �ٴ����ﰲ��.��ֹʱ��%Type := Null,
        strSQL = strSQL & "" & ZDate(dtEnd) & ","
        '  ���_In     �ٴ������.���%Type := Null,
        strSQL = strSQL & "" & intYear & ","
        '  �·�_In     �ٴ������.�·�%Type := Null,
        strSQL = strSQL & "" & intMonth & ","
        '  ����_In     �ٴ������.����%Type := Null,
        strSQL = strSQL & "" & ZVal(intWeek) & ","
        '  Ӧ�÷�Χ_In �ٴ������.Ӧ�÷�Χ%Type := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  ����id_In   �ٴ������.����id%Type := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  ��ע_In     �ٴ������.��ע%Type := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  ��Աid_In   ��Ա��.Id%Type := Null,
        strSQL = strSQL & "" & IIf(zlStr.IsHavePrivs(mstrPrivs, "���п���"), "NULL", UserInfo.ID) & ","
        '  ɾ������_In Number:=0
        strSQL = strSQL & "" & IIf(blnDeletePlan, 1, 0) & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Next
    If cllPlan.Count > 1 Then gcnOracle.CommitTrans
    
    'XX�³����ڵ㣺K2_���_�·�
    'XX�ܳ����ڵ㣺K3_���_�·�_����
    AddNewPlan = strKey
    Exit Function
TransErrHandler:
    If cllPlan.Count > 1 Then gcnOracle.RollbackTrans
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub chkShowAllPlan_Click()
    Err = 0: On Error GoTo errHandler
    
    lblPublishInfo.Visible = chkShowAllPlan.Value = vbUnchecked
    Screen.MousePointer = vbHourglass
    If mlng����ID = 0 And chkShowAllPlan.Value = vbUnchecked Then
        Set mrsPlanRecords = Nothing
    Else
        Set mrsPlanRecords = GetPlanRecords(mbytFun = 1, mlng����ID, Val(lblPublishInfo.Tag) = 1, _
            chkShowAllPlan.Value = vbChecked, mintYear, mintMonth, mdtStartDate, mdtEndDate)
    End If
    Call ExecuteFilter
'    Call ShowStopVisitPlan(vsfRegistPlan, mdtStartDate, mdtEndDate)
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Call mfrmMain.ActiveFormChange(Me)
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    
    Err = 0: On Error GoTo errHandler
    
    Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistPlan, Me.Name, "����")
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
    
    picSelectWeek.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, sccTitle.Width
    lineSplit.X1 = sccTitle.Left + 10
    lineSplit.Y1 = IIf(picSelectWeek.Visible, picSelectWeek.Top + picSelectWeek.Height, sccTitle.Top + sccTitle.Height - 10)
    lineSplit.X2 = sccTitle.Width
    lineSplit.Y2 = lineSplit.Y1
    With vsfRegistPlan
        .Left = sccTitle.Left + 10
        .Top = IIf(picSelectWeek.Visible, picSelectWeek.Top + picSelectWeek.Height + 20, sccTitle.Top + sccTitle.Height + 10)
        .Width = sccTitle.Width
        .Height = Me.ScaleHeight - .Top - 20
    End With
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
    
    objOut.Title.Text = Mid(sccTitle.Caption, InStr(sccTitle.Caption, ">") + 1) & "�嵥"
    If VSFlexGridCopyTo(vsfRegistPlan, vsfTemp, bytMode) = False Then Exit Sub
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

Private Sub Form_Unload(Cancel As Integer)
    Set mrsPlanRecords = Nothing
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistPlan, Me.Name, "����")
    Call SaveRegInFor(g˽��ģ��, Me.Name, "FindType", mintFindType)
End Sub

Private Sub optWeek_Click(index As Integer)
    Dim varDateRange As Variant, intWeek As Integer
    Dim dtStart As Date, dtEnd As Date
    
    Err = 0: On Error GoTo errHandler
    intWeek = index
    Screen.MousePointer = vbHourglass
    varDateRange = GetDateRange(mintYear, mintMonth, intWeek)
    dtStart = varDateRange(0): dtEnd = varDateRange(1)
    Call InitPlanGrid(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, dtStart, dtEnd, Val(lblPublishInfo.Tag) = 1)
    Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistPlan, Me.Name, "����")
    Call ShowHolidayToPlan(vsfRegistPlan, mdtStartDate, mdtEndDate)
    'ʹ�û�������
    Call ExecuteFilter
'    Call ShowStopVisitPlan(vsfRegistPlan, mdtStartDate, mdtEndDate)
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub sccTitle_GotFocus()
    On Error Resume Next
    If vsfRegistPlan.Visible And vsfRegistPlan.Enabled Then vsfRegistPlan.SetFocus
End Sub

Private Sub vsfRegistPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim cbrControl As CommandBarControl, lngCol As Long, dtCur As Date
    Dim strTemp As String, lng��¼ID As Long
    
    On Error Resume Next
    '�˵�����
    strTemp = Trim(vsfRegistPlan.Cell(flexcpData, 0, vsfRegistPlan.Col))
    If strTemp = "" Then Exit Sub
    If IsDate(strTemp) = False Then Exit Sub
    dtCur = CDate(strTemp)
    Set cbrControl = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_Edit_ApplyToDay, , True) 'Ӧ���ڵ���/˫��
    If Not cbrControl Is Nothing Then
        cbrControl.Caption = "Ӧ���ڡ�����" & IIf(Day(dtCur) Mod 2 = 0, "˫��", "����") & "��(&D)"
    End If
    Set cbrControl = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_Edit_ApplyToWeekDay, , True) 'Ӧ�������ڼ�
    If Not cbrControl Is Nothing Then
        cbrControl.Caption = "Ӧ���ڡ�����" & GetWeekName(Weekday(dtCur, vbMonday) - 1) & "��(&W)"
    End If
    
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
    Dim strCurItem As String, blnUpdate As Boolean
    Dim lngCol As Long, lngRow As Long
    Dim strSort As String
    Dim strTemp As String
    
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
            Call LoadPlanDataByRecordset(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, mrsPlanRecords, mbytFun, , True, _
                Val(lblPublishInfo.Tag) = 1, Format(mdtStartDate, "yyyy-mm-dd"), Format(mdtEndDate, "yyyy-mm-dd"))
'            Call ShowStopVisitPlan(vsfRegistPlan, mdtStartDate, mdtEndDate)
            Screen.MousePointer = vbDefault
        End If
    Else
        lng��ԴId = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ԴID))
        lng����ID = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_����ID))
        lngCol = GetPlanItemNameCol(vsfRegistPlan.Col)
        strCurItem = vsfRegistPlan.Cell(flexcpData, 0, lngCol)
        If lng��ԴId = 0 And lng����ID = 0 Then Exit Sub
        If IsDate(strCurItem) = False Then Exit Sub
        If Not (strCurItem >= mdtStartDate And strCurItem <= mdtEndDate) Then Exit Sub
        
        If chkShowAllPlan.Value = vbChecked Then
            lng����ID = 0: lng����ID = 0
            If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, vsfRegistPlan.Col)) = "" Then Exit Sub
            '�洢�ˡ�����ID,����ID��
            strTemp = vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col) + 2)
            If InStr(strTemp, ",") = 0 Then Exit Sub
            lng����ID = Val(Split(strTemp, ",")(0))
            lng����ID = Val(Split(strTemp, ",")(1))
        Else
            lng����ID = mlng����ID
        End If
        
        blnUpdate = zlStr.IsHavePrivs(mstrPrivs, "���ﰲ��") And Val(lblPublishInfo.Tag) = 0
        If zlStr.IsHavePrivs(mstrPrivs, "���п���") = False Then
            'û�С����п��ҡ�Ȩ��ʱ��ֻ�ܵ����������ٴ������Űࡱ�ĺ�Դ
            If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_�Ƿ��ٴ��Ű�)) = "" Then blnUpdate = False
        End If
        '��ʾ���а���ʱ��ֻ�ܲ鿴
        If chkShowAllPlan.Value = vbChecked Then blnUpdate = False
    
        If frmEdit.ShowMe(Me, mbytFun, IIf(blnUpdate, Fun_Update, Fun_View), lng����ID, lng��ԴId, lng����ID, strCurItem) Then
            If blnUpdate Then Call RefreshOneData
        End If
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

Private Function PublishPlan(ByVal lng����ID As Long, ByVal blnPublish As Boolean, _
    ByVal lng�����ܳ���ID As Long) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim obj�ϰ�ʱ�� As �ϰ�ʱ��
    Dim dtCurrent As Date, cll����ID  As New Collection, i As Integer
    Dim strPlanName As String
    
    Err = 0: On Error GoTo errHandler
    If blnPublish Then
        strSQL = "Select 1" & vbNewLine & _
                " From �ٴ����ﰲ�� A, �ٴ������¼ B, �ٴ������ C" & vbNewLine & _
                " Where a.Id = b.����id And a.����id = c.Id And c.�Ű෽ʽ In (1, 2)" & vbNewLine & _
                "       And c.Id In ([1], [2]) And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng�����ܳ���ID)
        If rsTemp.EOF Then
            MsgBox "��ǰ���������Ч�İ��ţ����ܷ�����", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select a.Id" & vbNewLine & _
                " From (Select ID, ��� || LPad(�·�, 2, '0') || ���� As ����" & vbNewLine & _
                "        From �ٴ������" & vbNewLine & _
                "        Where Nvl(�Ű෽ʽ, 0) = [2] And ������ Is Null And Id Not In ([1], [3])" & vbNewLine & _
                "              And Nvl(վ��,'-') = Nvl([4],'-')) A," & vbNewLine & _
                "      (Select ID, ��� || LPad(�·�, 2, '0') || ���� As ���� From �ٴ������ Where ID = [1]) B" & vbNewLine & _
                " Where a.���� < b.���� And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, mbytFun, lng�����ܳ���ID, gstrNodeNo)
        If Not rsTemp.EOF Then
            MsgBox "��ǰ�����ǰ�滹��δ������" & IIf(mbytFun = 1, "��", "��") & "����������Ƚ��䷢������ܷ����ó����", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select 1" & vbNewLine & _
                " From �ٴ������¼ A, �ٴ����ﰲ�� B" & vbNewLine & _
                " Where a.��Դid = b.��Դid And a.�������� Between b.��ʼʱ�� And b.��ֹʱ��" & vbNewLine & _
                "       And a.����ID <> b.Id And b.����id In ([1],[2]) And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng�����ܳ���ID)
        If Not rsTemp.EOF Then
            MsgBox "��ǰ������еĲ��ֺ�Դ�ڵ�ǰ��������Чʱ�䷶Χ���Ѿ�������Ч�İ��ţ����ܷ�����", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select Distinct d.����, d.����, e.վ��, b.��������, b.�ϰ�ʱ��, To_Char(c.��ʼʱ��, 'hh24:mi:ss') As ��ʼʱ��, " & vbNewLine & _
                "       To_Char(c.��ֹʱ��, 'hh24:mi:ss') As ��ֹʱ��" & vbNewLine & _
                " From �ٴ����ﰲ�� A, �ٴ������¼ B, �ٴ�������ſ��� C, �ٴ������Դ D, ���ű� E" & vbNewLine & _
                " Where a.Id = b.����id And b.Id = c.��¼id And a.��Դid = d.Id And d.����id = e.Id" & vbNewLine & _
                "       And c.��� = 1 And ����id In ([1],[2])" & vbNewLine & _
                " Order By " & IIf(gVisitPlan_ModulePara.byt����ȽϷ�ʽ = 0, "d.����", "Lpad(d.����,5,'0')")
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����ϰ�ʱ������ŷ�ʱ��ʱ���Ƿ�һ��", lng����ID, lng�����ܳ���ID)
        Do While Not rsTemp.EOF
            Set obj�ϰ�ʱ�� = GetWorkTimeRange(Nvl(rsTemp!�ϰ�ʱ��), Nvl(rsTemp!վ��), Nvl(rsTemp!����))
            If Format(obj�ϰ�ʱ��.��ʼʱ��, "hh:mm:00") <> Format(Nvl(rsTemp!��ʼʱ��), "hh:mm:00") Then
                If MsgBox("��ǰ������в��ַ�ʱ�εİ��Ų��Ǹ����ϰ�ʱ�ε�ʱ����зֶεģ��磺" & vbCrLf & _
                    "����Ϊ " & Nvl(rsTemp!����) & " ��" & Format(Nvl(rsTemp!��������), "yyyy-mm-dd") & " " & Nvl(rsTemp!�ϰ�ʱ��) & _
                    "[" & Format(obj�ϰ�ʱ��.��ʼʱ��, "hh:mm") & "-" & Format(obj�ϰ�ʱ��.����ʱ��, "hh:mm") & "]��" & _
                    "��һ�����ʱ��Ϊ[" & Format(Nvl(rsTemp!��ʼʱ��), "hh:mm") & "-" & Format(Nvl(rsTemp!��ֹʱ��), "hh:mm") & "])" & vbCrLf & vbCrLf & _
                    "�Ƿ���Ҫ����������", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                Exit Do
            End If
            rsTemp.MoveNext
        Loop
        
        strSQL = "Select Id, �������" & vbNewLine & _
                " From �ٴ������" & vbNewLine & _
                " Where Id In([1],[2]) And ����ʱ�� Is Null" & _
                " Order By ���,�·�,����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������Ϣ", lng����ID, lng�����ܳ���ID)
        If rsTemp.EOF Then
            MsgBox "��ǰ���������ѱ����˷�������ɾ������ˢ�����ݺ�鿴��", vbInformation, gstrSysName
            Exit Function
        End If
        Do While Not rsTemp.EOF
            cll����ID.Add Val(Nvl(rsTemp!ID))
            strPlanName = strPlanName & IIf(strPlanName <> "", "��", "") & "��" & Nvl(rsTemp!�������) & "��"
            rsTemp.MoveNext
        Loop
        
        If MsgBox("��ȷ��Ҫ����" & strPlanName & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        
        dtCurrent = zlDatabase.Currentdate
        On Error GoTo TransErrHandler
        
        Screen.MousePointer = vbHourglass
        If cll����ID.Count > 1 Then gcnOracle.BeginTrans
        For i = 1 To cll����ID.Count
            'Zl_�ٴ����ﰲ��_Publish
            strSQL = "Zl_�ٴ����ﰲ��_Publish("
            '  Id_In       �ٴ������.Id%Type,
            strSQL = strSQL & "" & cll����ID(i) & ","
            '  ������_In   �ٴ������.������%Type := Null,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '  ����ʱ��_In �ٴ������.����ʱ��%Type := Null,
            strSQL = strSQL & "" & ZDate(dtCurrent) & ")"
            '  ȡ������_In Number:=0
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Next
        If cll����ID.Count > 1 Then gcnOracle.CommitTrans
        Screen.MousePointer = vbDefault
    Else
        strSQL = "Select a.Id" & vbNewLine & _
                " From (Select ID, ��� || LPad(�·�, 2, '0') || ���� As ����" & vbNewLine & _
                "        From �ٴ������" & vbNewLine & _
                "        Where Nvl(�Ű෽ʽ, 0) = [2] And ������ Is Not Null And Id Not In ([1], [3])" & vbNewLine & _
                "              And Nvl(վ��,'-') = Nvl([4],'-')) A," & vbNewLine & _
                "      (Select ID, ��� || LPad(�·�, 2, '0') || ���� As ���� From �ٴ������ Where ID = [1]) B" & vbNewLine & _
                " Where a.���� > b.���� And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, mbytFun, lng�����ܳ���ID, gstrNodeNo)
        If Not rsTemp.EOF Then
            MsgBox "��ǰ������滹���ѷ�����" & IIf(mbytFun = 1, "��", "��") & "����������Ƚ���ȡ�����������ȡ�������ó����", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select 1" & vbNewLine & _
                " From ���˹Һż�¼ C, �ٴ������¼ A, �ٴ����ﰲ�� B" & vbNewLine & _
                " Where c.�����¼id = a.Id And a.����id = b.Id And b.����id In ([1],[2]) And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng�����ܳ���ID)
        If Not rsTemp.EOF Then
            MsgBox "��ǰ����������ܵİ����ѱ�ʹ�ã�������ȡ��������", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select Id, �������" & vbNewLine & _
                " From �ٴ������" & vbNewLine & _
                " Where Id In([1],[2]) And ����ʱ�� Is Not Null" & _
                " Order By ��� Desc,�·� Desc,���� Desc"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������Ϣ", lng����ID, lng�����ܳ���ID)
        If rsTemp.EOF Then
            MsgBox "��ǰ���������ѱ�����ȡ����������ɾ������ˢ�����ݺ�鿴��", vbInformation, gstrSysName
            Exit Function
        End If
        Do While Not rsTemp.EOF
            cll����ID.Add Val(Nvl(rsTemp!ID))
            strPlanName = "��" & Nvl(rsTemp!�������) & "��" & IIf(strPlanName <> "", "��", "") & strPlanName
            rsTemp.MoveNext
        Loop
        
        If MsgBox("��ȷ��Ҫȡ������" & strPlanName & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        dtCurrent = zlDatabase.Currentdate
        On Error GoTo TransErrHandler
        
        Screen.MousePointer = vbHourglass
        If cll����ID.Count > 1 Then gcnOracle.BeginTrans
        For i = 1 To cll����ID.Count
            'Zl_�ٴ����ﰲ��_Publish
            strSQL = "Zl_�ٴ����ﰲ��_Publish("
            '  Id_In       �ٴ������.Id%Type,
            strSQL = strSQL & "" & cll����ID(i) & ","
            '  ������_In   �ٴ������.������%Type := Null,
            strSQL = strSQL & "" & "NULL" & ","
            '  ����ʱ��_In �ٴ������.����ʱ��%Type := Null,
            strSQL = strSQL & "" & "NULL" & ","
            '  ȡ������_In Number:=0
            strSQL = strSQL & "" & 1 & ")"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Next
        If cll����ID.Count > 1 Then gcnOracle.CommitTrans
        Screen.MousePointer = vbDefault
    End If
    PublishPlan = True
    Exit Function
TransErrHandler:
    If cll����ID.Count > 1 Then gcnOracle.RollbackTrans
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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

Public Function PastPlan(ByVal lng����ID As Long, ByVal lngԭ����ID As Long, ByVal strԭ��Ŀ As String) As Long
    '���ܣ�ճ������
    '������
    Dim strSQL As String, strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim rsPlan As ADODB.Recordset, rsSignalSource As ADODB.Recordset
    Dim blnTran As Boolean, blnNoPlan As Boolean
    Dim lng����ID  As Long, lng��ԴId As Long, strApplyItem  As String
    Dim str���� As String
    
    Err = 0: On Error GoTo errHandler
    If lng����ID = 0 Then Exit Function
    If lngԭ����ID = 0 Then Exit Function
    If IsDate(strԭ��Ŀ) = False Then Exit Function
    
    With vsfRegistPlan
        lng����ID = Val(.TextMatrix(.Row, COL_����ID))
        lng��ԴId = Val(.TextMatrix(.Row, COL_��ԴID))
        strApplyItem = .Cell(flexcpData, 0, .Col)
        str���� = Trim(.TextMatrix(.Row, COL_����))
    End With
    
    If lng��ԴId = 0 Then Exit Function
    If IsDate(strApplyItem) = False Then Exit Function
    
    If lngԭ����ID = lng����ID And CDate(strԭ��Ŀ) = CDate(strApplyItem) Then
        MsgBox "��ǰ�����븴�ư�����ͬ������ճ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���ĳ���ϰ�ʱ���Ƿ������ڵ�ǰ��Դ
    strSQL = "Select b.����, a.�ϰ�ʱ��" & vbNewLine & _
            " From �ٴ������¼ A, �ٴ������Դ B" & vbNewLine & _
            " Where a.��ԴID = b.ID And a.����ID = [1] And a.�������� = [2] And a.�ϰ�ʱ�� Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngԭ����ID, CDate(strԭ��Ŀ))
    Do While Not rsTemp.EOF
        If GetWorkTimeRange(Nvl(rsTemp!�ϰ�ʱ��), gstrNodeNo, str����) Is Nothing Then
            MsgBox "�ϰ�ʱ�Ρ�" & Nvl(rsTemp!�ϰ�ʱ��) & "����������" & str���� & "�ţ�����ճ����", vbInformation, gstrSysName
            Exit Function
        End If
        rsTemp.MoveNext
    Loop
    
    If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, vsfRegistPlan.Col)) <> "" Then
        If MsgBox("��ճ�������ڵ�ǰ�Ѵ��ڳ��ﰲ�ţ�ճ�����ⲿ�ְ��Ž��ᱻ���ǣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    If lng����ID = 0 Then
        blnNoPlan = True
        
        '��鵱ǰ�����Ƿ������������������,һ����Դĳһ��İ���ֻ����һ�����������
        strSQL = "Select 1" & vbNewLine & _
                " From �ٴ������¼ A" & vbNewLine & _
                " Where a.�������� = [1] And A.��ԴId = [2] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(strApplyItem), lng��ԴId)
        If Not rsTemp.EOF Then
            MsgBox Format(strApplyItem, "yyyy-mm-dd") & " ��������������н����˰��ţ������ظ����ţ�", vbInformation, gstrSysName
            Exit Function
        End If
        
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
        strSQL = strSQL & "" & ZDate(mdtStartDate) & ","
        '��ֹʱ��_In     �ٴ����ﰲ��.��ֹʱ��%Type,
        strSQL = strSQL & "" & ZDate(mdtEndDate) & ","
        '����Ա����_In   �ٴ����ﰲ��.����Ա����%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '�Ǽ�ʱ��_In     �ٴ����ﰲ��.�Ǽ�ʱ��%Type
        strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ")"
    Else
        '��鵱ǰ�����Ƿ������������������,һ����Դĳһ��İ���ֻ����һ�����������
        strSQL = "Select 1" & vbNewLine & _
                " From �ٴ������¼ A" & vbNewLine & _
                " Where a.�������� = [1] And a.��ԴId = [2] And a.����id <> [3] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(strApplyItem), lng��ԴId, lng����ID)
        If Not rsTemp.EOF Then
            MsgBox Format(strApplyItem, "yyyy-mm-dd") & " ��������������н����˰��ţ������ظ����ţ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    blnTran = True
    gcnOracle.BeginTrans
        If blnNoPlan And strSQL <> "" Then
            zlDatabase.ExecuteProcedure strSQL, "��������"
        End If
        If ZlPlanApplyTo(1, lngԭ����ID, strԭ��Ŀ, lng����ID, strApplyItem) = False Then
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

Private Function ApplyToDay(ByVal lng����ID As Long, ByVal strCurDate As String) As Boolean
    '���ܣ�Ӧ���ڡ����е���/˫�ա�
    '������
    '   lng����ID ��Ӧ�õİ���ID
    '   dtCurDate ��Ӧ�õ�����
    Dim strApply As String, dtCur As Date
    Dim intDoubleDay As Integer
    Dim dtStart As Date, dtEnd As Date
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim lng��ԴId As Long
    
    Err = 0: On Error GoTo errHandler
    If lng����ID = 0 Or strCurDate = "" Then Exit Function
    If IsDate(strCurDate) = False Then Exit Function
    
    dtStart = CDate(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ʼʱ��))
    dtEnd = CDate(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ֹʱ��))
    lng��ԴId = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ԴID))
    
    '��ѯ��Դ�ǵ�ǰ��������õĳ����¼
    strSQL = "Select a.��������" & vbNewLine & _
            " From �ٴ������¼ A,�ٴ����ﰲ�� B,�ٴ������ C" & vbNewLine & _
            " Where a.����ID=b.ID And b.����ID<>[1] And a.��ԴID=[2] And a.�������� Between [3] And [4]" & vbNewLine & _
            "       And c.ID=b.����ID And Nvl(c.�Ű෽ʽ,0) In (1,2)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Դ�������˵ĳ����¼", mlng����ID, lng��ԴId, dtStart, dtEnd)
    
    intDoubleDay = Day(strCurDate) Mod 2 '���ջ���˫��
    dtCur = dtStart
    Do While DateDiff("d", dtCur, dtEnd) >= 0
        If DateDiff("d", strCurDate, dtCur) <> 0 And (Day(dtCur) Mod 2) = intDoubleDay Then
            rsTemp.Filter = "��������=#" & Format(dtCur, "yyyy-mm-dd") & "#"
            If rsTemp.RecordCount = 0 Then
                strApply = strApply & "|" & Format(dtCur, "yyyy-mm-dd")
            End If
        End If
        dtCur = DateAdd("d", 1, dtCur)
    Loop
    If strApply <> "" Then strApply = Mid(strApply, 2)
    
    If strApply = "" Then Exit Function
    If CheckExistRecord(lng��ԴId, strApply) Then
        If MsgBox("ע�⣺" & vbCrLf & _
                  "      ���ֱ�Ӧ�õ����ڵ�ǰ�Ѵ��ڳ��ﰲ�ţ�Ӧ�ú��ⲿ�ְ��Ž��ᱻ���ǣ��Ƿ���ҪӦ�ã�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    ApplyToDay = ZlPlanApplyTo(1, lng����ID, Format(strCurDate, "yyyy-mm-dd"), lng����ID, strApply)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
        
Private Function ApplyToWeekDay(ByVal lng����ID As Long, ByVal strCurDate As String) As Boolean
    '���ܣ�Ӧ���ڡ��������ڼ���
    '������
    '   lng����ID ��Ӧ�õİ���ID
    '   dtCurDate ��Ӧ�õ�����
    Dim strApply As String, dtCur As Date
    Dim intWeekDay As Integer
    Dim dtStart As Date, dtEnd As Date
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim lng��ԴId As Long
    
    Err = 0: On Error GoTo errHandler
    If lng����ID = 0 Or strCurDate = "" Then Exit Function
    If IsDate(strCurDate) = False Then Exit Function
    
    dtStart = CDate(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ʼʱ��))
    dtEnd = CDate(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ֹʱ��))
    lng��ԴId = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ԴID))
    
    '��ѯ��Դ�ǵ�ǰ��������õĳ����¼
    strSQL = "Select a.��������" & vbNewLine & _
            " From �ٴ������¼ A,�ٴ����ﰲ�� B,�ٴ������ C" & vbNewLine & _
            " Where a.����ID=b.ID And b.����ID<>[1] And a.��ԴID=[2] And a.�������� Between [3] And [4]" & vbNewLine & _
            "       And c.ID=b.����ID And Nvl(c.�Ű෽ʽ,0) In (1,2)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Դ�������˵ĳ����¼", mlng����ID, lng��ԴId, dtStart, dtEnd)
    
    intWeekDay = Weekday(strCurDate, vbMonday) '���ڼ�
    dtCur = dtStart
    Do While DateDiff("d", dtCur, dtEnd) >= 0
        If DateDiff("d", strCurDate, dtCur) <> 0 And Weekday(dtCur, vbMonday) = intWeekDay Then
            rsTemp.Filter = "��������=#" & Format(dtCur, "yyyy-mm-dd") & "#"
            If rsTemp.RecordCount = 0 Then
                strApply = strApply & "|" & Format(dtCur, "yyyy-mm-dd")
            End If
        End If
        dtCur = DateAdd("d", 1, dtCur)
    Loop
    If strApply <> "" Then strApply = Mid(strApply, 2)
    
    If strApply = "" Then Exit Function
    If CheckExistRecord(lng��ԴId, strApply) Then
        If MsgBox("ע�⣺" & vbCrLf & _
                  "      ���ֱ�Ӧ�õ����ڵ�ǰ�Ѵ��ڳ��ﰲ�ţ�Ӧ�ú��ⲿ�ְ��Ž��ᱻ���ǣ��Ƿ���ҪӦ�ã�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    ApplyToWeekDay = ZlPlanApplyTo(1, lng����ID, Format(strCurDate, "yyyy-mm-dd"), lng����ID, strApply)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function NextNewPlanByPlan(ByVal lngԭ����ID As Long, Optional ByVal blnMonth As Boolean) As Boolean
    '�������г���������°���
    Dim strSQL As String, rsTemp As ADODB.Recordset, lng����ID As Long
    Dim intYear As Integer, intMonth As Integer, intWeek As Integer
    Dim dtStart As Date, dtEnd As Date
    Dim strName As String, strKey As String, blnDeletePlan As Boolean
    Dim cllPlan As Collection, i As Integer
    Dim dtCurrent As Date
    
    Err = 0: On Error GoTo errHandler
    Set cllPlan = GetNewPlanInfo(Me, mstrPrivs, blnMonth, strKey, blnDeletePlan)
    If cllPlan Is Nothing Then Exit Function
    If cllPlan.Count = 0 Then Exit Function
    
    dtCurrent = zlDatabase.Currentdate
    On Error GoTo TransErrHandler
        
    Screen.MousePointer = vbHourglass
    If cllPlan.Count > 1 Then gcnOracle.BeginTrans
    For i = 1 To cllPlan.Count
        'Array(���,�·�,����,��ʼ����,��������)
        intYear = cllPlan(i)(0)
        intMonth = cllPlan(i)(1)
        intWeek = cllPlan(i)(2)
        dtStart = cllPlan(i)(3)
        dtEnd = cllPlan(i)(4)
    
        'ȷ�������ID
        strSQL = "Select ID From �ٴ������ Where �Ű෽ʽ = [1] And ��� = [2] And �·� = [3]" & _
            IIf(blnMonth, "", " And ���� = [4]") & " And Nvl(վ��,'-') = Nvl([5],'-')"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(blnMonth, 1, 2), intYear, intMonth, intWeek, gstrNodeNo)
        If rsTemp.EOF Then
            lng����ID = zlDatabase.GetNextId("�ٴ������")
        Else
            lng����ID = Val(Nvl(rsTemp!ID))
        End If
        
        strName = intYear & "��" & intMonth & "��"
        If Not blnMonth Then strName = strName & "��" & intWeek & "��"
        strName = strName & "�����"
        
        'zl_�ٴ������_Addbyrecord(
        strSQL = "zl_�ٴ������_Addbyrecord("
        'ԭ����Id_In         �ٴ������.Id%Type,
        strSQL = strSQL & "" & lngԭ����ID & ","
        '�³���id_In      �ٴ������.Id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '�Ű෽ʽ_In   �ٴ������.�Ű෽ʽ%Type,
        strSQL = strSQL & "" & IIf(blnMonth, 1, 2) & ","
        '�������_In   �ٴ������.�������%Type,
        strSQL = strSQL & "'" & strName & "',"
        '���_In     �ٴ������.���%Type,
        strSQL = strSQL & "" & intYear & ","
        '�·�_In     �ٴ������.�·�%Type,
        strSQL = strSQL & "" & intMonth & ","
        '����_In     �ٴ������.����%Type := Null,
        strSQL = strSQL & "" & ZVal(intWeek) & ","
        '��ʼʱ��_In �ٴ����ﰲ��.��ʼʱ��%Type := Null,
        strSQL = strSQL & "" & ZDate(dtStart) & ","
        '��ֹʱ��_In �ٴ����ﰲ��.��ֹʱ��%Type := Null,
        strSQL = strSQL & "" & ZDate(dtEnd) & ","
        '����Ա_In   �ٴ����ﰲ��.����Ա����%Type := Null,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '�Ǽ�ʱ��_In �ٴ����ﰲ��.�Ǽ�ʱ��%Type := Null,
        strSQL = strSQL & "" & ZDate(dtCurrent) & ","
        'վ��_In       ���ű�.վ��%Type,
        strSQL = strSQL & "'" & gstrNodeNo & "',"
        'ȫԺ��Դ����վ��_In ���ű�.վ��%Type,
        strSQL = strSQL & "'" & gVisitPlan_ModulePara.str��Դά��վ�� & "',"
        '��Աid_In     ��Ա��.Id%Type := Null,
        strSQL = strSQL & "" & IIf(HavePrivs(mstrPrivs, "���п���"), "NULL", UserInfo.ID) & ","
        'ɾ������_In Number:=0
        strSQL = strSQL & "" & IIf(blnDeletePlan, 1, 0) & ")"
        
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Next
    If cllPlan.Count > 1 Then gcnOracle.CommitTrans
    
    'XX�³����ڵ㣺K2_���_�·�
    'XX�ܳ����ڵ㣺K3_���_�·�_����
    Call mfrmMain.NodeChanged(strKey)
    NextNewPlanByPlan = True
    
    Screen.MousePointer = vbDefault
    Exit Function
TransErrHandler:
    If cllPlan.Count > 1 Then gcnOracle.RollbackTrans
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPlanRecords(ByVal blnMonth As Boolean, Optional ByVal lng����ID As Long, Optional ByVal blnPublished As Boolean, _
    Optional ByVal blnMonthAllPlan As Boolean, Optional intYear As Integer, Optional intMonth As Integer, _
    Optional ByVal dtStartDate As Date, Optional ByVal dtEndDate As Date, Optional ByVal lng�����ܳ���ID As Long) As ADODB.Recordset
    '���ܣ���ȡ���ż�¼
    '������
    '   blnMonth - �Ƿ����Ű�
    '   lng����ID   - ����ID
    '   blnPublished- �Ƿ��ѷ���
    '   blnMonthAllPlan - �Ƿ���ʾ�������а���
    '   lng�����ܳ���ID - �����ܿ��µ���һ�������
    Dim strSQL As String, strSqlSub As String
    Dim strWhere As String, str��Ч��Դ As String
    Dim str�Ƿ���Ч As String
    Dim str������� As String

    Err = 0: On Error GoTo errHandler
    str������� = IIf(gVisitPlan_ModulePara.byt����ȽϷ�ʽ = 0, "a.����", "Lpad(a.����,5,'0')")
    strSqlSub = "       " & str������� & " As �������, a.Id As ��Դid, a.����, a.����, Nvl(a.�Ƿ񽨲���, 0) As �Ƿ񽨲���,a.ԤԼ����, a.����Ƶ��," & vbNewLine & _
                "       Decode(a.���տ���״̬, 1, '����ԤԼ', 2, '��ֹԤԼ', 3, '�ܽڼ������ÿ���', '���ϰ�') As ���տ���״̬," & vbNewLine & _
                "       Nvl(a.�Ƿ��ٴ��Ű�, 0) As �Ƿ��ٴ��Ű�, Decode(a.�Ű෽ʽ, 1, '�����Ű�', 2, '�����Ű�', '�̶��Ű�') As �Ű෽ʽ," & vbNewLine & _
                "       f.���� As ����, f.���� As ���Ҽ���,Nvl(a.�Ƿ���ջ���, 0) As �Ƿ���ջ���," & vbNewLine
    
    'û��"���п���"Ȩ�޵Ĳ���Աֻ�ܲ����Լ��������ҵĺ�Դ
    If HavePrivs(mstrPrivs, "���п���") = False Then
        strWhere = "      And Exists (Select 1 From ������Ա Where ����id = a.����id And ��Աid = [3])"
    End If
    
    str�Ƿ���Ч = "Decode(Sign(b.��ֹʱ��- Trunc(Sysdate)), -1, 0, 1)*Decode(b.���ʱ��, NULL, 0, 1) As �Ƿ���Ч,"
    
    If blnMonth And blnMonthAllPlan Then
        '��ѯ�������а���
        '�Ȳ��ҳ���ID
        strSQL = "Select m.Id" & vbNewLine & _
                " From �ٴ������ M, �ٴ������ N" & vbNewLine & _
                " Where m.�Ű෽ʽ In(1, 2) And m.��� = [1] And m.�·� = [2] And Nvl(m.վ��,'-') = Nvl([4],'-')"
        
        strSQL = "Select " & str�Ƿ���Ч & vbNewLine & _
                "        b.����id, b.Id As ����id, e.���� As �շ���Ŀ, b.ҽ������, g.���� As ҽ������, g.רҵ����ְ�� as ҽ��ְ��,h.��ʶ��, " & vbNewLine & strSqlSub & _
                "        c.Id As ��¼id, c.��������, c.�ϰ�ʱ��, c.�޺���, c.��Լ��, c.�ѹ���, c.��Լ��, c.ԤԼ���� As ԤԼ���Ʒ�ʽ," & vbNewLine & _
                "        c.�Ƿ���ʱ����, c.ͣ�￪ʼʱ��, c.ͣ����ֹʱ��, c.ͣ��ԭ��, c.����ҽ������, c.�Ƿ�����, b.��ʼʱ��, b.��ֹʱ��" & vbNewLine & _
                " From �ٴ������Դ A, �ٴ����ﰲ�� B, �ٴ������¼ C, �շ���ĿĿ¼ E, ���ű� F, ��Ա�� G,רҵ����ְ�� H" & vbNewLine & _
                " Where a.Id = b.��Դid And b.Id = c.����id And a.����id = f.Id And b.��Ŀid = e.Id And b.ҽ��ID = g.ID(+) and g.רҵ����ְ��=h.����(+)" & vbNewLine & _
                "       And Nvl(a.�Ƿ�ɾ��, 0) = 0 And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & vbNewLine & _
                "       And (e.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or e.����ʱ�� Is Null)" & vbNewLine & _
                "       And (f.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or f.����ʱ�� Is Null)" & vbNewLine & _
                "       And (g.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or g.����ʱ�� Is Null)" & vbNewLine & _
                "       And b.����id In (" & strSQL & ")" & strWhere & vbNewLine & _
                " Order By " & str������� & ", ��������, �ϰ�ʱ��"
        Set GetPlanRecords = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Ű���Ϣ", intYear, intMonth, UserInfo.ID, gstrNodeNo)
        Exit Function
    End If
    
    '��ѯ�°��Ż��ܰ���
    If blnPublished Then
        strSQL = "Select " & str�Ƿ���Ч & vbNewLine & _
                "        b.����id, b.Id As ����id, e.���� As �շ���Ŀ, b.ҽ������, g.���� As ҽ������, g.רҵ����ְ�� as ҽ��ְ��,h.��ʶ��, " & vbNewLine & strSqlSub & _
                "        c.Id As ��¼id, c.��������, c.�ϰ�ʱ��, c.�޺���, c.��Լ��, c.�ѹ���, c.��Լ��, c.ԤԼ���� As ԤԼ���Ʒ�ʽ," & vbNewLine & _
                "        c.�Ƿ���ʱ����, c.ͣ�￪ʼʱ��, c.ͣ����ֹʱ��, c.ͣ��ԭ��, c.����ҽ������, c.�Ƿ�����, b.��ʼʱ��, b.��ֹʱ��" & vbNewLine & _
                " From �ٴ������Դ A, �ٴ����ﰲ�� B, �ٴ������¼ C, �շ���ĿĿ¼ E, ���ű� F, ��Ա�� G,רҵ����ְ�� H" & vbNewLine & _
                " Where a.Id = b.��Դid And b.Id = c.����id(+) And a.����id = f.Id And b.��Ŀid = e.Id And b.ҽ��ID = g.Id(+) and g.רҵ����ְ��=h.����(+)" & vbNewLine & _
                "       And (b.����id = [1] Or b.����id = [7])" & strWhere & vbNewLine & _
                "       And Nvl(a.�Ƿ�ɾ��, 0) = 0 And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & vbNewLine & _
                "       And (e.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or e.����ʱ�� Is Null)" & vbNewLine & _
                "       And (f.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or f.����ʱ�� Is Null)" & vbNewLine & _
                "       And (g.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or g.����ʱ�� Is Null)" & vbNewLine & _
                " Order By " & str������� & ", ��������, �ϰ�ʱ��"
    Else
        If blnMonth Then
            str��Ч��Դ = " And a.�Ű෽ʽ = [2]"
        Else
            '��ǰ�ѵ���Ϊ�����Ű�,���Ǳ������������Ű࣬����ʣ�µĲ��ֽ��������ܽ����Ű�
            'ͬʱ,��ǰ���������ʱ�䷶Χ�ڲ��������Ű�
            str��Ч��Դ = " And (a.�Ű෽ʽ = [2] And Not Exists (Select 1" & vbNewLine & _
                        "           From �ٴ������¼ O, �ٴ����ﰲ�� P, �ٴ������ Q" & vbNewLine & _
                        "           Where o.����id = p.Id And p.����id = q.Id And p.��Դid+0 = a.Id" & vbNewLine & _
                        "               And o.�������� Between [6] And Last_Day([4]) And q.�Ű෽ʽ = 1)" & vbNewLine & _
                        "       Or a.�Ű෽ʽ = 1 And Exists (Select 1" & vbNewLine & _
                        "           From �ٴ������¼ O, �ٴ����ﰲ�� P, �ٴ������ Q" & vbNewLine & _
                        "           Where o.����id = p.Id And p.����id = q.Id And p.��Դid+0 = a.Id" & vbNewLine & _
                        "               And o.�������� Between [6] And Last_Day([4]) And q.�Ű෽ʽ = 2))" & vbNewLine
        End If
        '��û���������ŵĺ�Դ��Ҫ���޳����¼ ��
        str��Ч��Դ = str��Ч��Դ & vbNewLine & _
                    "       And Not Exists" & vbNewLine & _
                    "           (Select 1" & vbNewLine & _
                    "            From �ٴ������¼ P" & vbNewLine & _
                    "            Where p.��Դid+0 = a.Id And p.�������� Between [4] And [5])" & vbNewLine
        'δ����ʱ�������к�Դȱʡ��ȡ����
        strSQL = "Select " & str�Ƿ���Ч & vbNewLine & _
                "        b.����id, b.Id As ����id, " & vbNewLine & strSqlSub & _
                "        Decode(b.ID,Null,e.����,m.����) As �շ���Ŀ, Decode(b.ID,Null,a.ҽ������,b.ҽ������) As ҽ������," & vbNewLine & _
                "        Decode(b.ID,Null,g.����,n.����) As ҽ������, Decode(b.ID,Null,g.רҵ����ְ��,n.רҵ����ְ��) as ҽ��ְ��," & vbNewLine & _
                "        Decode(b.ID,Null,i.��ʶ��,j.��ʶ��) As ��ʶ�� ," & vbNewLine & _
                "        c.Id As ��¼id, c.��������, c.�ϰ�ʱ��, c.�޺���, c.��Լ��, b.��ʼʱ��, b.��ֹʱ��, " & vbNewLine & _
                "        c.�ѹ���, c.��Լ��, c.ԤԼ���� As ԤԼ���Ʒ�ʽ, c.�Ƿ���ʱ����, c.ͣ�￪ʼʱ��, c.ͣ����ֹʱ��, c.ͣ��ԭ��, c.����ҽ������, c.�Ƿ�����" & vbNewLine & _
                " From �ٴ������Դ A, " & vbNewLine & _
                "      (Select ����id, ID, ��Դid, ��Ŀid, ҽ��ID, ҽ������, ��ʼʱ��, ��ֹʱ��, ���ʱ��" & vbNewLine & _
                "        From �ٴ����ﰲ�� Where (����id = [1] Or ����id = [7])) B," & vbNewLine & _
                "      �ٴ������¼ C, �շ���ĿĿ¼ E, ���ű� F, ��Ա�� G, �շ���ĿĿ¼ M, ��Ա�� N,רҵ����ְ�� I,רҵ����ְ�� J" & vbNewLine & _
                " Where a.Id = b.��Դid(+) And b.Id = c.����id(+) And a.����id = f.Id" & vbNewLine & _
                "       And a.��Ŀid = e.Id And a.ҽ��ID = g.ID(+) And b.��Ŀid = m.Id(+) And b.ҽ��ID = n.ID(+)" & vbNewLine & _
                "       And g.רҵ����ְ��=i.����(+) And n.רҵ����ְ��=j.����(+)" & vbNewLine & _
                "       And Nvl(a.�Ƿ�ɾ��, 0) = 0 And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & vbNewLine & _
                "       And (e.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or e.����ʱ�� Is Null)" & vbNewLine & _
                "       And (f.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or f.����ʱ�� Is Null)" & vbNewLine & _
                "       And (g.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or g.����ʱ�� Is Null)" & vbNewLine & _
                "       And (b.Id Is Not Null Or (b.Id Is Null " & str��Ч��Դ & "))" & strWhere & vbNewLine & _
                "       And Nvl(Nvl(f.վ��,[9]),Nvl([8],'-')) = Nvl([8],'-')" & vbNewLine & _
                " Order By " & str������� & ", c.��������, c.�ϰ�ʱ��"
    End If
    Set GetPlanRecords = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Ű���Ϣ", lng����ID, IIf(blnMonth, 1, 2), _
        UserInfo.ID, dtStartDate, dtEndDate, CDate(Format(dtStartDate, "yyyy-mm-01")), lng�����ܳ���ID, _
        gstrNodeNo, gVisitPlan_ModulePara.str��Դά��վ��)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtFind_KeyPress(index As Integer, KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        Call ExecuteFilter
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

Private Function SaveAsTemplet(ByVal lng����ID As Long, Optional ByVal blnMonth As Boolean) As Boolean
    '���Ϊģ��
    Dim obj���ﰲ�� As New ���ﰲ��, frmPlanInfoEdit As New frmClinicPlanInfoEdit
    Dim strSQL As String, strAddVisitTableSQL As String
    Dim rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    If lng����ID = 0 Then Exit Function
    '����Ƿ�����Ч�İ���
    If ExistsPlanOnVisitTable(IIf(blnMonth, 1, 2), lng����ID, IIf(zlStr.IsHavePrivs(mstrPrivs, "���п���"), 0, UserInfo.ID)) = False Then
        MsgBox "��ǰ�����������Ч�İ��ţ��������Ϊģ�壡", vbInformation, gstrSysName
        Exit Function
    End If
    
    obj���ﰲ��.�Ű෽ʽ = 3 '�Ű෽ʽ��0-�̶��Ű�;1-�����Ű�;2-�����Ű�;3-ģ��
    obj���ﰲ��.ģ������ = IIf(mbytFun = 1, 2, 0) '0-���Ű�ģ�壬1-���ǰ����Ű�����Ű�ģ�壬2-�����Ű�����Ű�ģ��
    If frmPlanInfoEdit.ShowMe(Me, mlngModule, 1, obj���ﰲ��, False, True) = False Then Exit Function
    
    obj���ﰲ��.����ID = zlDatabase.GetNextId("�ٴ������")
    'Zl_�ٴ������_Totemplet(
    strSQL = "Zl_�ٴ������_Totemplet("
    '  ����id_In   �ٴ������.Id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ģ��id_In   �ٴ������.Id%Type,
    strSQL = strSQL & "" & obj���ﰲ��.����ID & ","
    '  �������_In �ٴ������.�������%Type,
    strSQL = strSQL & "'" & obj���ﰲ��.������� & "',"
    '  Ӧ�÷�Χ_In �ٴ������.Ӧ�÷�Χ%Type,
    strSQL = strSQL & "" & obj���ﰲ��.Ӧ�÷�Χ & ","
    '  ����id_In   �ٴ������.����id%Type,
    strSQL = strSQL & "" & ZVal(obj���ﰲ��.����ID) & ","
    '  ��ע_In     �ٴ������.��ע%Type,
    strSQL = strSQL & "'" & obj���ﰲ��.��ע & "',"
    '  ����Ա_In   �ٴ����ﰲ��.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ����ʱ��_In �ٴ����ﰲ��.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ","
    '  վ��_In     ���ű�.վ��%Type,
    strSQL = strSQL & IIf(gstrNodeNo = "", "NULL", "'" & gstrNodeNo & "'") & ","
    '  ��Աid_In   ��Ա��.Id%Type := Null
    strSQL = strSQL & "" & IIf(zlStr.IsHavePrivs(mstrPrivs, "���п���"), "NULL", UserInfo.ID) & ")"
    
    Screen.MousePointer = vbHourglass
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    '�����ģ��ڵ㣺K0_����ID
    Call mfrmMain.NodeChanged("K0_" & obj���ﰲ��.����ID)
    SaveAsTemplet = True
    
    Screen.MousePointer = vbDefault
    Exit Function
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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
