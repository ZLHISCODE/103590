VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmRegistPlan 
   AutoRedraw      =   -1  'True
   Caption         =   "�ҺŰ��Ź���"
   ClientHeight    =   9435
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12030
   Icon            =   "frmRegistPlan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   4050
      ScaleHeight     =   2520
      ScaleWidth      =   5070
      TabIndex        =   0
      Top             =   1470
      Width           =   5070
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4290
         _Version        =   589884
         _ExtentX        =   7567
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistPlan.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistPlan.frx":049E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   9075
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRegistPlan.frx":07F2
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16140
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
      Left            =   0
      Top             =   1170
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmRegistPlan.frx":1086
      Left            =   480
      Top             =   1365
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmRegistPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mArrFilter As Variant, mlngModule As Long, mstrPrivs As String, mblnDisStop As Boolean, mblnDisDel As Boolean
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar
Private Enum mPgIndex
    Pg_��ǰ��Ч�ű� = 1
    Pg_�ƻ����źű� = 2
End Enum
Private Const ID_PANE_SEARCH = 1
Private Const ID_PANE_Page = 2

Private mPanSearch As Pane
Private mblnUnload  As Boolean
Private mfrm���źű� As frmRegistPlanPlan
Private WithEvents mfrm��Ч�ű� As frmRegistPlanList
Attribute mfrm��Ч�ű�.VB_VarHelpID = -1
Private WithEvents mfrmFilter As frmRegistPlanFilter
Attribute mfrmFilter.VB_VarHelpID = -1
Private mfrmUnitReg  As frmCooperateUnitsReg
Private mfrmUnitRegPlan As frmCooperateUnitsRegPlan

Private mblnFirst As Boolean

Private mbln�Զ�Ĭ����Լ�� As Boolean '45519
Private mblnԤԼ�����ڽ�ֹɾ�� As Boolean
Private Sub zlRptPrint(bytMode As Byte)
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If Val(tbPage.Selected.Tag) = mPgIndex.Pg_��ǰ��Ч�ű� Then
        mfrm��Ч�ű�.zlRptPrint bytMode
    Else
        mfrm���źű�.zlRptPrint bytMode
    End If
End Sub
Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-09-09 15:04:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, strKey As String
    If mfrmFilter Is Nothing Then Set mfrmFilter = New frmRegistPlanFilter
    Load mfrmFilter
    '��ʼ�� ���� �Ƿ���ʾͣ�ð���
    mfrmFilter.ShowStop = mblnDisStop
    mfrmFilter.ShowDel = mblnDisDel
    Set mArrFilter = mfrmFilter.GetCondition
    
    With dkpMan
        .ImageList = imlPaneIcons
        Set mPanSearch = .CreatePane(ID_PANE_SEARCH, 300, 100, DockLeftOf, Nothing)
        mPanSearch.Title = "��������": mPanSearch.Options = PaneNoCloseable
         Set objPane = .CreatePane(ID_PANE_Page, 400, 400, DockRightOf, mPanSearch)
        objPane.Title = ""
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picList.Hwnd
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    dkpMan.RecalcLayout: DoEvents
    zlRestoreDockPanceToReg Me, dkpMan, "����"
    Call GetRegInFor(g˽��ģ��, Me.Name, "����", strKey)
    If Val(strKey) = 1 Then mPanSearch.Hide
    mPanSearch.MinTrackSize.Width = 230: mPanSearch.MaxTrackSize.Width = 230
       
End Function
Private Sub zlRefreshData()
    zlCommFun.ShowFlash "����װ������,���Ժ�..."
    Call InitData
    Set mArrFilter = mfrmFilter.GetCondition
    Call mfrm��Ч�ű�.zlRefreshData(mArrFilter)
    Call mfrm���źű�.zlRefreshData(mArrFilter)
    If Val(tbPage.Selected.Tag) = mPgIndex.Pg_��ǰ��Ч�ű� Then
        Call mfrm��Ч�ű�.zlActtion
    Else
        Call mfrm���źű�.zlActtion
    End If
    zlCommFun.StopFlash
End Sub
Private Sub zlPlanManager(ByVal bytFun As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����/ȡ��/���/ȡ����ˣ��ƻ�����
    '���:bytFun-(0-����,1-ȡ��,2-���,3-ȡ�����,4-����
    '����:���˺�
    '����:2009-09-15 17:16:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, lng�ƻ�ID As Long
    frmRegistPlanArrange.�Զ�Ĭ����Լ�� = mbln�Զ�Ĭ����Լ��
    Select Case bytFun
    Case 0  '���Ӽƻ�����
        If Val(tbPage.Selected.Tag) = mPgIndex.Pg_�ƻ����źű� Then Exit Sub
        lngID = mfrm��Ч�ű�.zlGet����ID
        If lngID = 0 Then Exit Sub
        If frmRegistPlanArrange.ShowCard(Me, mlngModule, mstrPrivs, ed_�ƻ�����, lngID, "") = False Then Exit Sub
    Case 5  '�޸ļƻ�
        If Val(tbPage.Selected.Tag) <> mPgIndex.Pg_�ƻ����źű� Then
            lngID = mfrm��Ч�ű�.zlGet����ID
            lng�ƻ�ID = mfrm��Ч�ű�.zlGet����ID(True)
        Else
            lngID = mfrm���źű�.zlGet����ID(False)
            lng�ƻ�ID = mfrm���źű�.zlGet����ID(True)
        End If
        If lng�ƻ�ID = 0 Then Exit Sub
        If frmRegistPlanArrange.ShowCard(Me, mlngModule, mstrPrivs, Ed_�����޸�, lngID, lng�ƻ�ID) = False Then Exit Sub
         Call mfrm��Ч�ű�.ReloadTimePlan(True)
    Case 1  'ȡ���ƻ�����
        If Val(tbPage.Selected.Tag) <> mPgIndex.Pg_�ƻ����źű� Then
            lngID = mfrm��Ч�ű�.zlGet����ID
            lng�ƻ�ID = mfrm��Ч�ű�.zlGet����ID(True)
        Else
            lngID = mfrm���źű�.zlGet����ID(False)
            lng�ƻ�ID = mfrm���źű�.zlGet����ID(True)
        End If
        If lng�ƻ�ID = 0 Then Exit Sub
        If frmRegistPlanArrange.ShowCard(Me, mlngModule, mstrPrivs, Ed_����ɾ��, lngID, lng�ƻ�ID) = False Then Exit Sub
    Case 2   '���
        If Val(tbPage.Selected.Tag) <> mPgIndex.Pg_�ƻ����źű� Then
            lngID = mfrm��Ч�ű�.zlGet����ID
            lng�ƻ�ID = mfrm��Ч�ű�.zlGet����ID(True)
        Else
            lngID = mfrm���źű�.zlGet����ID(False)
            lng�ƻ�ID = mfrm���źű�.zlGet����ID(True)
        End If
        If lng�ƻ�ID = 0 Then Exit Sub
        If frmRegistPlanArrange.ShowCard(Me, mlngModule, mstrPrivs, Ed_�������, lngID, lng�ƻ�ID) = False Then Exit Sub
    Case 3   'ȡ��
        If Val(tbPage.Selected.Tag) <> mPgIndex.Pg_�ƻ����źű� Then
            lngID = mfrm��Ч�ű�.zlGet����ID
            lng�ƻ�ID = mfrm��Ч�ű�.zlGet����ID(True)
        Else
            lngID = mfrm���źű�.zlGet����ID(False)
            lng�ƻ�ID = mfrm���źű�.zlGet����ID(True)
        End If
        If lng�ƻ�ID = 0 Then Exit Sub
        If CheckPlanBooking(lng�ƻ�ID) Then
            If MsgBox("��ǰ�ƻ��Ѿ�����ԤԼ�Һŵ�����ȷ��Ҫȡ�������", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        If frmRegistPlanArrange.ShowCard(Me, mlngModule, mstrPrivs, Ed_����ȡ��, lngID, lng�ƻ�ID) = False Then Exit Sub
    Case 4  '����
    
        If Val(tbPage.Selected.Tag) <> mPgIndex.Pg_�ƻ����źű� Then
            lngID = mfrm��Ч�ű�.zlGet����ID
            lng�ƻ�ID = mfrm��Ч�ű�.zlGet����ID(True)
        Else
            lngID = mfrm���źű�.zlGet����ID(False)
            lng�ƻ�ID = mfrm���źű�.zlGet����ID(True)
        End If
        If lng�ƻ�ID = 0 Then Exit Sub
        If frmRegistPlanArrange.ShowCard(Me, mlngModule, mstrPrivs, ed_���Ų���, lngID, lng�ƻ�ID) = False Then Exit Sub
        Exit Sub
    Case Else
    End Select
    
    If Val(tbPage.Selected.Tag) <> mPgIndex.Pg_�ƻ����źű� Then
        zlCommFun.ShowFlash "����װ������,���Ժ�..."
        mfrm��Ч�ű�.zlRefreshOlnyPlanData
        mfrm��Ч�ű�.Tag = "1"
        mfrm��Ч�ű�.zlActtion
        mfrm���źű�.Tag = ""
        zlCommFun.StopFlash
    Else
        zlCommFun.ShowFlash "����װ������,���Ժ�..."
        mfrm��Ч�ű�.Tag = ""
        mfrm���źű�.Tag = "1"
        mfrm���źű�.zlRefreshData mArrFilter
        mfrm���źű�.zlActtion
        zlCommFun.StopFlash
    End If
End Sub
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '------------------------------------
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_Edit_NewItem   '����
        Call zlAddItem
    Case conMenu_Edit_Modify    '�޸�
        Call zlModifyItem
    Case conMenu_Edit_Delete 'ɾ������
        Call zlDeleteItem
    Case conMenu_View_Refresh   'ˢ��
        Call zlRefreshData
    Case conMenu_View_ShowStoped '��ʾͣ�ð���
        mblnDisStop = IIf(mblnDisStop, False, True)
        Call zlDatabase.SetPara("��ʾͣ�ð���", IIf(mblnDisStop, "1", "0"), glngSys, mlngModule)
        mfrmFilter.ShowStop = mblnDisStop
        Call zlRefreshData
    Case conMenu_View_ShowDel '��ʾɾ������
        mblnDisDel = IIf(mblnDisDel, False, True)
        Call zlDatabase.SetPara("��ʾɾ������", IIf(mblnDisDel, "1", "0"), glngSys, mlngModule)
        mfrmFilter.ShowDel = mblnDisDel
        Call zlRefreshData
    Case conMenu_Edit_AllStartNO     'ȫ��������ſ���
        Call BatchSet(1)
    Case conMenu_Edit_AllStopNO      'ȫ��ȡ����ſ���
        Call BatchSet(0)
    Case conMenu_Edit_Reuse       ' "���ð���(&I)")
        Call zlStopAndResume(False)
    Case conMenu_Edit_Stop          ' "ͣ�ð���(&T)"):
        Call zlStopAndResume(True)
    Case conMenu_Edit_StopPlanTimes '����ͣ�üƻ�
        Call zlStopPlanTimes
    Case conMenu_Edit_ClearStopPlan '�������ͣ�üƻ�
        Call zlClearStopPlanTimes
    Case conMenu_Manage_Bespeak 'ʱ�������
        '*****************
        '�ϰ�ʱ�������
        '*****************
          frmSplitTime.Show 1, Me
    Case ComMenu_Edit_AutoDefaultLimitAppointment 'Ĭ����Լ��
         Control.Checked = Not Control.Checked
         mbln�Զ�Ĭ����Լ�� = Control.Checked
         Call zlDatabase.SetPara("�Զ�Ĭ����Լ��", IIf(mbln�Զ�Ĭ����Լ��, 1, 0), glngSys, mlngModule)
    Case comMenu_Edit_SetDateSegment
'        '*****************
'        '�ҺŰ���ʱ�������
'        '*****************
'        '�����:51429
'        If Control.Caption = "����ʱ������" Then
'            Call zlSetDateSegment
'        Else
'            Call zlSetPlanDateSegment
'        End If
    Case conMenu_Edit_SetPlanDateSeqment
          '*********************
          '�ҺŰ��żƻ�ʱ�������
          '*********************
          Call zlSetPlanDateSegment
    Case comMenu_Edit_UnitRegModify  '������λ���ſ���
        zlExecuteUnitReg
    Case ComMenu_Edit_UnitRegArrangeModify '������λ�ƻ�����
        Call zlExecuteUnitReg(True)
        
    Case conMenu_Edit_PlanAdd   '�ƻ�����
        Call zlPlanManager(0)  '0-����,1-ȡ��,2-���,3-ȡ�����,4-����,5-�޸�
    Case conMenu_Edit_PlanModify   '�޸ļƻ�����
        Call zlPlanManager(5)  '0-����,1-ȡ��,2-���,3-ȡ�����,4-����,5-�޸�
    Case conMenu_Edit_PlanDelete   'ȡ������
        Call zlPlanManager(1)  '0-����,1-ȡ��,2-���,3-ȡ�����,4-����,5-�޸�
    Case conMenu_Edit_PlanVerify      '���
        Call zlPlanManager(2)  '0-����,1-ȡ��,2-���,3-ȡ�����,4-����,5-�޸�
    Case conMenu_Edit_PlanCancel   'ȡ�����
        Call zlPlanManager(3)  '0-����,1-ȡ��,2-���,3-ȡ�����,4-����,5-�޸�
    Case conMenu_File_Exit: Unload Me
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_View_StatusBar
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
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.Hwnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.Hwnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            'ִ�з�������ǰģ��ı���
            If Val(Me.tbPage.Selected.Tag) = Pg_��ǰ��Ч�ű� Then
                Call mfrm��Ч�ű�.zlCallCustomReprot(Me, Val(Split(Control.Parameter, ",")(0)), Trim(Split(Control.Parameter, ",")(1)))
            Else
                Call mfrm���źű�.zlCallCustomReprot(Me, Val(Split(Control.Parameter, ",")(0)), Trim(Split(Control.Parameter, ",")(1)))
            End If
        End If
    End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
 

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, lngID As Long, blnEnabled As Boolean
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_��ǰ��Ч�ű� Then
            Control.Enabled = mfrm��Ч�ű�.zlGet����ID <> 0
        Else
            Control.Enabled = mfrm���źű�.zlGet����ID <> 0
        End If
    Case conMenu_Edit_NewItem '����
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_��ǰ��Ч�ű� Then
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
        Else
            Control.Visible = False
        End If
        Control.Enabled = Control.Visible
    Case conMenu_Edit_AllStartNO     'ȫ��������ſ���
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_AllStopNO      'ȫ��ȡ����ſ���
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_StopPlanTimes
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_��ǰ��Ч�ű� Then
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
            Control.Enabled = Control.Visible And Not mfrm��Ч�ű�.zlIsStopPlan And mfrm��Ч�ű�.zlGet����ID <> 0
        Else
            Control.Visible = False
        End If
    Case conMenu_Edit_Stop  'ͣ�ð���
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_��ǰ��Ч�ű� And zlStr.IsHavePrivs(mstrPrivs, "ͣ�ð���") Then
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
            Control.Enabled = Control.Visible And Not mfrm��Ч�ű�.zlIsStopPlan And mfrm��Ч�ű�.zlGet����ID <> 0
        Else
            Control.Visible = False
        End If
    Case conMenu_Edit_Reuse '���ð���
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_��ǰ��Ч�ű� And zlStr.IsHavePrivs(mstrPrivs, "���ð���") Then
                Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
                Control.Enabled = Control.Visible And mfrm��Ч�ű�.zlIsStopPlan And mfrm��Ч�ű�.zlGet����ID <> 0
        Else
            Control.Visible = False
        End If
    Case conMenu_Edit_PlanAdd   '�ƻ�����
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_��ǰ��Ч�ű� Then
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "���Ӽƻ�")
            Control.Enabled = mfrm��Ч�ű�.zlGet����ID <> 0
        Else
            Control.Visible = False
        End If
        Control.Enabled = Control.Visible And Control.Enabled
        If Control.Enabled Then
            Control.Enabled = Not mfrm��Ч�ű�.zlGet����ͣ��
        End If
    Case conMenu_Edit_PlanModify    '�޸ļƻ�
        If zlStr.IsHavePrivs(mstrPrivs, "�޸�����") = True Then
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�޸ļƻ�")
            If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_��ǰ��Ч�ű� Then
                'zlPlanStatus() '0-�����ڼƻ�����,1-δ���,2-�Ѿ����,3-�Ѿ���Ч\
                lngID = mfrm��Ч�ű�.zlPlanStatus
                Control.Enabled = lngID <> 0
            Else
                'zlPlanStatus() '0-�����ڼƻ�����,1-δ���,2-�Ѿ����,3-�Ѿ���Ч\
                lngID = mfrm���źű�.zlPlanStatus
                Control.Enabled = lngID <> 0
            End If
            Control.Enabled = Control.Visible And Control.Enabled
        Else
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�޸ļƻ�")
            If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_��ǰ��Ч�ű� Then
                'zlPlanStatus() '0-�����ڼƻ�����,1-δ���,2-�Ѿ����,3-�Ѿ���Ч\
                lngID = mfrm��Ч�ű�.zlPlanStatus
                Control.Enabled = lngID = 1
            Else
                'zlPlanStatus() '0-�����ڼƻ�����,1-δ���,2-�Ѿ����,3-�Ѿ���Ч\
                lngID = mfrm���źű�.zlPlanStatus
                Control.Enabled = lngID = 1
            End If
            Control.Enabled = Control.Visible And Control.Enabled
        End If
    Case conMenu_Edit_PlanDelete, conMenu_Edit_SetPlanDateSeqment, ComMenu_Edit_UnitRegArrangeModify  'ɾ���ƻ�
        If Control.ID = conMenu_Edit_PlanModify Or Control.ID = conMenu_Edit_SetPlanDateSeqment Then
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�޸ļƻ�")
        Else
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "ɾ���ƻ�")
        End If
        If Control.ID = ComMenu_Edit_UnitRegArrangeModify Then
            '�������
             Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�Һź�����λ����")
        End If
        
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_��ǰ��Ч�ű� Then
            'zlPlanStatus() '0-�����ڼƻ�����,1-δ���,2-�Ѿ����,3-�Ѿ���Ч\
            lngID = mfrm��Ч�ű�.zlPlanStatus
            Control.Enabled = lngID = 1
        Else
            'zlPlanStatus() '0-�����ڼƻ�����,1-δ���,2-�Ѿ����,3-�Ѿ���Ч\
            lngID = mfrm���źű�.zlPlanStatus
            Control.Enabled = lngID = 1
        End If
        Control.Enabled = Control.Visible And Control.Enabled
        If Control.ID = conMenu_Edit_SetPlanDateSeqment Then
            Control.Enabled = mfrm��Ч�ű�.zlHaveDatePlan(True) And Control.Enabled
        End If
    Case conMenu_Edit_ClearStopPlan '�������ͣ�üƻ�
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�������ͣ�üƻ�")
        
    Case conMenu_Edit_PlanVerify      '���
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�ƻ����")
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_��ǰ��Ч�ű� Then
            'zlPlanStatus() '0-�����ڼƻ�����,1-δ���,2-�Ѿ����,3-�Ѿ���Ч\
            lngID = mfrm��Ч�ű�.zlPlanStatus
            Control.Enabled = lngID = 1
        Else
            'zlPlanStatus() '0-�����ڼƻ�����,1-δ���,2-�Ѿ����,3-�Ѿ���Ч\
            lngID = mfrm���źű�.zlPlanStatus
            Control.Enabled = lngID = 1
        End If
        Control.Enabled = Control.Visible And Control.Enabled
    Case conMenu_Edit_PlanCancel   'ȡ�����
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "ȡ�����")
         If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_��ǰ��Ч�ű� Then
            'zlPlanStatus() '0-�����ڼƻ�����,1-δ���,2-�Ѿ����,3-�Ѿ���Ч\
            lngID = mfrm��Ч�ű�.zlPlanStatus
            Control.Enabled = (lngID <> 3 And lngID = 2)
        Else
            'zlPlanStatus() '0-�����ڼƻ�����,1-δ���,2-�Ѿ����,3-�Ѿ���Ч\
            lngID = mfrm���źű�.zlPlanStatus
            Control.Enabled = (lngID <> 3 And lngID = 2)
        End If
        Control.Enabled = Control.Visible And Control.Enabled
    Case conMenu_Edit_Modify
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_��ǰ��Ч�ű� Then
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
            'zlPlanStatus() '0-�����ڼƻ�����,1-δ���,2-�Ѿ����,3-�Ѿ���Ч\
            lngID = mfrm��Ч�ű�.zlPlanStatus
            Control.Enabled = mfrm��Ч�ű�.zlGet����ID <> 0 And Not mfrm��Ч�ű�.zlIsStopPlan
        Else
            Control.Visible = False
        End If
        Control.Enabled = Control.Enabled And Control.Visible
    Case conMenu_Edit_Delete, comMenu_Edit_UnitRegModify  ', comMenu_Edit_SetDateSegment
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_��ǰ��Ч�ű� Then
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
            'zlPlanStatus() '0-�����ڼƻ�����,1-δ���,2-�Ѿ����,3-�Ѿ���Ч\
            lngID = mfrm��Ч�ű�.zlPlanStatus
            Control.Enabled = mfrm��Ч�ű�.zlGet����ID <> 0 And lngID = 0 And Not mfrm��Ч�ű�.zlIsStopPlan
        Else
            Control.Visible = False
        End If
        If Control.ID = comMenu_Edit_UnitRegModify Then
             Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�Һź�����λ����")
        End If
        Control.Enabled = Control.Enabled And Control.Visible
'        If Control.ID = comMenu_Edit_SetDateSegment Then
'             '�����:51427
'            If mfrm��Ч�ű�.�Ƿ�ѡ�мƻ��б� = False Then
'                Control.Enabled = Not mfrm��Ч�ű�.Have�ƻ�
'                Control.Caption = "����ʱ������"
'            Else
'                If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_��ǰ��Ч�ű� Then
'                'zlPlanStatus() '0-�����ڼƻ�����,1-δ���,2-�Ѿ����,3-�Ѿ���Ч\
'                        lngID = mfrm��Ч�ű�.zlPlanStatus
'                        Control.Enabled = lngID = 1
'                Else
'                'zlPlanStatus() '0-�����ڼƻ�����,1-δ���,2-�Ѿ����,3-�Ѿ���Ч\
'                        lngID = mfrm���źű�.zlPlanStatus
'                        Control.Enabled = lngID = 1
'                End If
'                Control.Enabled = Control.Visible And Control.Enabled
'                Control.Enabled = mfrm��Ч�ű�.zlHaveDatePlan(True) And Control.Enabled
'                Control.Caption = "�ƻ�ʱ������"
'            End If
'        End If
    Case ComMenu_Edit_AutoDefaultLimitAppointment  '�Զ�Ĭ����Լ��
         Control.Checked = mbln�Զ�Ĭ����Լ��
      '��ʾͣ�ð���
    Case conMenu_View_ShowStoped
         Control.Checked = mblnDisStop
         '��ʾɾ������
    Case conMenu_View_ShowDel
         Control.Checked = mblnDisDel
    Case conMenu_Manage_Bespeak 'ʱ�������
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "ʱ�������")
        Control.Enabled = Control.Visible
    Case conMenu_View_Refresh
                
    End Select
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
'    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case ID_PANE_SEARCH
        Item.Handle = mfrmFilter.Hwnd
    Case ID_PANE_Page
        Item.Handle = picList.Hwnd
    End Select
End Sub
Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2009-09-09 11:01:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo Errhand:
    
 
    Set mfrm��Ч�ű� = New frmRegistPlanList
    Set ObjItem = tbPage.InsertItem(mPgIndex.Pg_��ǰ��Ч�ű�, "��ǰ��Ч�ű�", mfrm��Ч�ű�.Hwnd, 0)
    ObjItem.Tag = mPgIndex.Pg_��ǰ��Ч�ű�
    Call mfrm��Ч�ű�.UpdatePara(mfrmFilter.chkShowExpiredPlan.Value = 1)
    
    Set mfrm���źű� = New frmRegistPlanPlan
    Set ObjItem = tbPage.InsertItem(mPgIndex.Pg_�ƻ����źű�, "�ƻ����źű�", mfrm���źű�.Hwnd, 0)
    ObjItem.Tag = mPgIndex.Pg_�ƻ����źű�

     With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
    stbThis.Top = Me.ScaleHeight - Me.stbThis.Height
End Sub
'Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
'    Top = IIf(cbr.Visible, cbr.Height, 0)
'    Bottom = IIf(stbThis.Visible, stbThis.Height, 0)
'End Sub
'Private Sub Form_Resize()
'    Dim cbrH As Long '������ռ�ø߶�
'    Dim staH As Long '״̬��ռ�ø߶�
'    Dim i As Integer, lngW As Long
'
'    On Error Resume Next
'    If WindowState = 1 Then Exit Sub
'    '����ؼ���Ⱥ͸߶�
'    cbrH = IIf(cbr.Visible, cbr.Height, 0)
'    staH = IIf(stbThis.Visible, stbThis.Height, 0)
'    With mshPlan
'        .Left = Me.ScaleLeft
'        .Top = Me.ScaleTop + cbrH
'        .Width = Me.ScaleWidth
'        .Height = Me.ScaleHeight - cbrH - staH
'    End With
'End Sub


Private Sub Form_Activate()
    Dim strKey As String
    If mblnUnload Then Unload Me: Exit Sub
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
'    Form_Resize
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    
    mstrPrivs = gstrPrivs: mlngModule = glngModul
    mblnFirst = True
    '��ȡ���� �Ƿ���ʾͣ�ð���
    mblnDisStop = Val(zlDatabase.GetPara("��ʾͣ�ð���", glngSys, mlngModule, 0)) = 1
    '��ȡ���� �Ƿ���ʾɾ������
    mblnDisDel = Val(zlDatabase.GetPara("��ʾɾ������", glngSys, mlngModule, 0)) = 1
    mbln�Զ�Ĭ����Լ�� = Val(zlDatabase.GetPara("�Զ�Ĭ����Լ��", glngSys, mlngModule, 0)) = 1
    '46639ԤԼ������
    mblnԤԼ�����ڽ�ֹɾ�� = Val(zlDatabase.GetPara("ԤԼ�����ڽ�ֹɾ��", glngSys, mlngModule)) = 1
    Call InitData
    Call zlCommFun.SetWindowsInTaskBar(Me.Hwnd, False)
    Call zlDefCommandBars
    Call InitPanel
    Call InitPage
    '��ȡ����
    Call zlRefreshData
    'Ȩ�޴���
    'Call Ȩ�޿���
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
End Sub
Private Sub InitData()
    Dim strSQL As String
    '���¼ƻ�:���ƻ�����
    strSQL = "Zl_�ҺŰ���_Autoupdate"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    zlSaveDockPanceToReg Me, dkpMan, "����"
    Call SaveRegInFor(g˽��ģ��, Me.Name, "����", IIf(mPanSearch.Hidden, 1, 0))
    If Not mfrmFilter Is Nothing Then Unload mfrmFilter
    If Not mfrm���źű� Is Nothing Then Unload mfrm���źű�
    If Not mfrm��Ч�ű� Is Nothing Then Unload mfrm��Ч�ű�
    
    Set mfrmFilter = Nothing
    Set mfrm���źű� = Nothing
    Set mfrm��Ч�ű� = Nothing
    SaveWinState Me, App.ProductName
End Sub
Private Sub mfrmFilter_zlRefreshCon(ByVal ArrFilter As Variant)
    Set mArrFilter = ArrFilter
    Call mfrm��Ч�ű�.UpdatePara(mfrmFilter.chkShowExpiredPlan.Value = 1)
    Call mfrm��Ч�ű�.ReloagUnitRegPlan
    '���������˸ı�
    Select Case Val(tbPage.Selected.Tag)
    Case mPgIndex.Pg_��ǰ��Ч�ű�
        Call mfrm��Ч�ű�.zlRefreshData(ArrFilter)
    Case mPgIndex.Pg_�ƻ����źű�
        Call mfrm���źű�.zlRefreshData(ArrFilter)
    Case Else
    End Select
End Sub
Private Sub mfrm��Ч�ű�_zlPopuMenu(intType As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = 2 And intType = 0) Then Exit Sub
    Dim objPopup As CommandBarPopup
    Set objPopup = cbsThis.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
End Sub

Private Sub zlSetDateSegment()
    '************************
    '�ҺŰ���ʱ������
    '************************
    Dim lng����ID       As Long
    lng����ID = mfrm��Ч�ű�.zlGet����ID(False)
    If ExistsBooking(lng����ID) Then
         If MsgBox("�úű����ԤԼ�Һŵ�,�Ƿ��޸�ʱ��?", vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
            Exit Sub
         End If
    End If
    If frmRegistPlanDatSet.ShowMe(lng����ID, Edit) Then
         mfrm��Ч�ű�.ReloadTimePlan
    End If
End Sub

Private Sub zlSetPlanDateSegment()
     '************************
    '�ҺŰ��żƻ�ʱ������
    '************************
    Dim lng�ƻ�ID       As Long
    If Val(tbPage.Selected.Tag) <> mPgIndex.Pg_�ƻ����źű� Then
        lng�ƻ�ID = mfrm��Ч�ű�.zlGet����ID(True)
    Else
        lng�ƻ�ID = mfrm���źű�.zlGet����ID(True)
    End If
   ' lng�ƻ�ID = mfrm��Ч�ű�.zlGet����ID(True)
    If frmRegistPlanPlanDatSet.ShowMe(lng�ƻ�ID, Edit) Then
         mfrm��Ч�ű�.ReloadTimePlan (True)
    End If
    
End Sub
Private Sub zlExecuteUnitReg(Optional ByVal blnPlan As Boolean = False)
     '************************
    '
    '************************
    Dim lng����ID       As Long
    Dim lng�ƻ�ID       As Long
    If blnPlan = False Then
        lng����ID = mfrm��Ч�ű�.zlGet����ID(False)
        If ExistsBooking(lng����ID) Then
            Call MsgBox("�úű����ԤԼ�Һŵ�,�����޸ĺ�����λԤԼ�ŷ��䣡", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName)
            Exit Sub
        End If
        If Not mfrmUnitReg Is Nothing Then Set mfrmUnitReg = Nothing
        Set mfrmUnitReg = New frmCooperateUnitsReg
        If mfrmUnitReg.zlShowMe(lng����ID, mlngModule, mstrPrivs) Then              'ˢ��
           mfrm��Ч�ű�.zl_ReLoadUnitReg
        End If
        Set mfrmUnitReg = Nothing
    Else
         '************************
        '
        '************************
        
        If Val(tbPage.Selected.Tag) <> mPgIndex.Pg_�ƻ����źű� Then
            lng�ƻ�ID = mfrm��Ч�ű�.zlGet����ID(True)
        Else
            lng�ƻ�ID = mfrm���źű�.zlGet����ID(True)
        End If
        If Not mfrmUnitRegPlan Is Nothing Then Set mfrmUnitRegPlan = Nothing
            Set mfrmUnitRegPlan = New frmCooperateUnitsRegPlan
            If mfrmUnitRegPlan.zlShowMe(lng�ƻ�ID, mlngModule, mstrPrivs) Then
                mfrm��Ч�ű�.ReloagUnitRegPlan
                  
            End If
            Set mfrmUnitRegPlan = Nothing
    End If
End Sub

Private Sub zlAddItem()
    frmRegistPlanEdit.�Զ�Ĭ����Լ�� = mbln�Զ�Ĭ����Լ��
    If frmRegistPlanEdit.ShowEdit(Me, edt_����, mlngModule, mstrPrivs, 0, mfrmFilter.zlGet����ID) = False Then Exit Sub
    mfrm��Ч�ű�.zlRefreshData (mArrFilter)
End Sub
Private Sub BatchSet(bytFun As Byte)
    Dim strSQL As String
    Dim i As Long
        
    If MsgBox("��ȷ��Ҫ�������޺Ż���Լ�ĺű�" & IIf(bytFun = 1, "����", "ȡ��") & "��ſ�����?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then
        Exit Sub
    End If
    
    On Error GoTo errH
    strSQL = "Zl_�ҺŰ���_��ſ���(" & bytFun & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    mfrm��Ч�ű�.zlRefreshData mArrFilter
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function CheckExistsBooking(str�ű� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ���ű��Ƿ����ԤԼ�Һŵ�
    '���:str�ű�-�ű�
    '����:����,����true,���򷵻�False
    '����:���˺�
    '����:2009-09-15 10:32:02
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select /*+ Rule*/ Min(����ʱ��) ʱ��" & vbNewLine & _
            "From ������ü�¼" & vbNewLine & _
            "Where ��¼���� = 4 And ��¼״̬ In (0, 1) And ���㵥λ = [1] And ����ʱ�� > �Ǽ�ʱ��"
    If gintԤԼ���� = 0 Then
        strSQL = strSQL & " And ����ʱ�� > Sysdate"
    Else
        strSQL = strSQL & " And ����ʱ�� Between Sysdate And Sysdate+" & gintԤԼ����
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str�ű�)
    
    CheckExistsBooking = Not IsNull(rsTmp!ʱ��)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckPlanBooking(lng�ƻ�ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ���ƻ��Ƿ����ԤԼ�Һŵ�
    '���:str�ű�-�ű�
    '����:����,����true,���򷵻�False
    '����:���˺�
    '����:2009-09-15 10:32:02
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 1" & vbNewLine & _
            " From ���˹Һż�¼ A, �ҺŰ��żƻ� B" & vbNewLine & _
            " Where b.���� = a.�ű� And a.��¼״̬ = 1 And a.����ʱ�� Between b.��Чʱ�� + 0 And b.ʧЧʱ�� And b.���ʱ�� Is Not Null And b.Id = [1]"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ƻ�ID)
    
    CheckPlanBooking = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub zlDeleteItem()
    'ɾ���ű�
    mfrmFilter.ShowDel = mblnDisDel
    Set mArrFilter = mfrmFilter.GetCondition
    If mfrm��Ч�ű�.zlExecuteDeleteList(mblnԤԼ�����ڽ�ֹɾ��) = False Then Exit Sub
    mfrm��Ч�ű�.Tag = ""
End Sub
Private Sub zlModifyItem()
    frmRegistPlanEdit.�Զ�Ĭ����Լ�� = mbln�Զ�Ĭ����Լ�� '45519
    If mfrm��Ч�ű�.zlExecuteModifyList(Me) = False Then Exit Sub
    mfrm��Ч�ű�.Tag = ""
    mfrm��Ч�ű�.ReloadTimePlan
End Sub
 
 
Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub

 
Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Not mfrmFilter Is Nothing Then
        mfrmFilter.zlblnShowPlanCon = Val(Item.Tag) = mPgIndex.Pg_�ƻ����źű�
    End If
    zlCommFun.ShowFlash "����װ������,���Ժ�..."
   If Val(tbPage.Selected.Tag) = mPgIndex.Pg_��ǰ��Ч�ű� Then
        If mfrm��Ч�ű�.Tag = "" Then
            mfrm��Ч�ű�.zlRefreshOlnyPlanData
            mfrm��Ч�ű�.Tag = "1"
        End If
        Call mfrm��Ч�ű�.zlActtion
    Else
        If mfrm���źű�.Tag = "" Then
            mfrm���źű�.zlRefreshData mArrFilter
            mfrm���źű�.Tag = "1"
        End If
        Call mfrm���źű�.zlActtion
    End If
    zlCommFun.StopFlash
End Sub

Public Function zlDefCommandBars() As Boolean
    '----------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/1/9
    '----------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
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
        'Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "���Ӱ���(&A)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸İ���(&M)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ������(&D)")
'        '�����:51156
'        Set mcbrControl = .Add(xtpControlButton, comMenu_Edit_SetDateSegment, "����ʱ������"): mcbrControl.IconId = 3063
        Set mcbrControl = .Add(xtpControlButton, comMenu_Edit_UnitRegModify, "������λ���ſ���"): mcbrControl.IconId = 3813
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_AllStartNO, "ȫ��������ſ���(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_AllStopNO, "ȫ��ȡ����ſ���(&U)"):
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "���ð���(&I)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "ͣ�ð���(&T)"):
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_StopPlanTimes, "����ͣ�üƻ�(&Q)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClearStopPlan, "�������ͣ�üƻ�(&W)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanAdd, "���Ӽƻ�(&N)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanModify, "�޸ļƻ�(&G)")
        'Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_SetPlanDateSeqment, "����ʱ��"): mcbrControl.IconId = 3063:
        Set mcbrControl = .Add(xtpControlButton, ComMenu_Edit_UnitRegArrangeModify, "������λ�ƻ�����"): mcbrControl.IconId = 3813
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanDelete, "ɾ���ƻ�(&R)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanVerify, "��˼ƻ�(&V)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanCancel, "ȡ�����(&C)")
        '�����:51156
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Bespeak, "�ϰ�ʱ������(&T)"): mcbrControl.IconId = 3038: mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, ComMenu_Edit_AutoDefaultLimitAppointment, "�Զ�Ĭ����Լ��")
        mcbrControl.Checked = mbln�Զ�Ĭ����Լ��
        
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
         '��ʾͣ�ð���
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ShowStoped, "��ʾͣ��(&D)")
        mcbrControl.Checked = mblnDisStop
         '����: 45525
         '��ʾɾ������
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ShowDel, "��ʾɾ��(&Z)")
        mcbrControl.Checked = mblnDisDel
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
    
    '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("T"), comMenu_Edit_SetDateSegment
        .Add FCONTROL, Asc("V"), conMenu_Edit_PlanVerify
         
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
        
    End With
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
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
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "���Ӱ���"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸İ���")
        '�����:51156
        'Set mcbrControl = .Add(xtpControlButton, comMenu_Edit_SetDateSegment, "����ʱ������"): mcbrControl.IconId = 3063:
       ' Set mcbrControl = .Add(xtpControlButton, comMenu_Edit_UnitRegModify, "������λ���ſ���"): mcbrControl.IconId = 3813
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ������")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "���ð���"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "ͣ�ð���"):
   
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanAdd, "���Ӽƻ�"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanModify, "�޸ļƻ�"): mcbrControl.BeginGroup = True
        '�����:51156
        'Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_SetPlanDateSeqment, "����ʱ��"): mcbrControl.IconId = 3063:
       ' Set mcbrControl = .Add(xtpControlButton, ComMenu_Edit_UnitRegArrangeModify, "������λ�ƻ�����"): mcbrControl.IconId = 3813
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanDelete, "ɾ���ƻ�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanVerify, "�ƻ����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanCancel, "ȡ�����")
        '�����:51156
        'Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Bespeak, "ʱ���"): mcbrControl.IconId = 3063: mcbrControl.BeginGroup = True
        

        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
     zlDefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub zlStopAndResume(Optional blnStop As Boolean = True)
    'ɾ���ű�
    mfrmFilter.ShowStop = mblnDisStop
    Set mArrFilter = mfrmFilter.GetCondition
    If mfrm��Ч�ű�.zlStopAndResume(blnStop) = False Then Exit Sub
    mfrm��Ч�ű�.Tag = ""
End Sub
Private Sub zlStopPlanTimes()
    If mfrm��Ч�ű�.zlStopPlanTimes() = False Then Exit Sub
    mfrm��Ч�ű�.Tag = ""
End Sub
Private Sub zlClearStopPlanTimes()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ͣ�üƻ�
    '����:���˺�
    '����:2010-09-09 14:42:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrm��Ч�ű�.zlClearStopPlanTimes() = False Then Exit Sub
    mfrm��Ч�ű�.Tag = ""
End Sub
 
 
Private Function ExistsBooking(ByVal lng����ID As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ���ű��Ƿ����ԤԼ�Һŵ�
    '���:str�ű�-�ű�
    '����:����,����true,���򷵻�False
    '����:
    '����:2012-04-26 10:32:02
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select min(A.����ʱ��) as ʱ��  From ���˹Һż�¼ A, �ҺŰ��� B "
    strSQL = strSQL & vbCrLf & " Where A.�ű� = B.���� "
    strSQL = strSQL & vbCrLf & "       And ��¼״̬ = 1 and b.id=[1] And ����ʱ�� > �Ǽ�ʱ�� "
    If gintԤԼ���� = 0 Then
        strSQL = strSQL & " And A.����ʱ�� > Sysdate"
    Else
        strSQL = strSQL & " And A.����ʱ�� Between Sysdate And Sysdate+" & gintԤԼ����
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    ExistsBooking = Not IsNull(rsTmp!ʱ��)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

