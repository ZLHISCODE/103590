VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmClinicPlanTempletManage 
   BorderStyle     =   0  'None
   Caption         =   "����ģ�����"
   ClientHeight    =   7260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   9360
      MaxLength       =   100
      TabIndex        =   4
      Top             =   930
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfRegistPlan 
      Height          =   2085
      Left            =   390
      TabIndex        =   1
      Top             =   1050
      Width           =   2535
      _cx             =   4471
      _cy             =   3678
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
      FormatString    =   $"frmClinicPlanTempletManage.frx":0000
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
         Picture         =   "frmClinicPlanTempletManage.frx":0075
         ScaleHeight     =   135
         ScaleWidth      =   150
         TabIndex        =   2
         Top             =   90
         Width           =   150
      End
   End
   Begin VB.Label lblPlanInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ�÷�Χ����������(�����ڿ�)  ��ע������ط�����ٸԷ���ط������涨���򹦷�"
      Height          =   180
      Left            =   3960
      TabIndex        =   3
      Top             =   150
      Width           =   6840
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   10845
      _Version        =   589884
      _ExtentX        =   19129
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "���ﰲ��>����ģ��"
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
      Height          =   6915
      Left            =   0
      Top             =   0
      Width           =   11595
   End
End
Attribute VB_Name = "frmClinicPlanTempletManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar�ؼ�
Private mlngModule As Long
Private mstrPrivs As String

Private mbytFun As Byte '1-��ģ��,2-��ģ��
Private mlng����ID As Long
Private mrsPlanRecords As ADODB.Recordset
Private mintFindType As Integer

Private mstrOldSelRangePlan As String '��¼ѡ���������򣬸�ʽ"��ʼ��|������|��ʼ��|������"�����ڻָ�����ɫ

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
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&R)", cbrControl.Index + 1): cbrControl.BeginGroup = True
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddTemplet, "����ģ��(&T)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�ģ��(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��ģ��(&D)")

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyPlanItem, "��������(&D)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyUnitRegist, "����ԤԼ�Һſ���(&U)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AllStartNO, "ȫ��������ſ���(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AllStopNO, "ȫ��ȡ����ſ���(&T)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearAllPlan, "�����ǰ��Դ����(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearAll, "������к�Դ����(&A)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextMonthNewPlan, "�����³����(&N)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextWeekNewPlan, "�����ܳ����(&W)"): cbrControl.BeginGroup = True
    End With

    '�鿴�˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) 'ˢ����ǰ(���ʱע�ⷴ��)
'        Set cbrControl = .Add(xtpControlButton,conMenu_View_Notify,"ˢ������(&B)",cbrControl.Index)
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddTemplet, "����ģ��", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�ģ��", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��ģ��", cbrControl.index + 1)
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyPlanItem, "��������", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyUnitRegist, "ԤԼ�Һſ���", cbrControl.index + 1)
        cbrControl.ToolTipText = "����ԤԼ�Һſ���"

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextMonthNewPlan, "�����³����", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextWeekNewPlan, "�����ܳ����", cbrControl.index + 1): cbrControl.BeginGroup = True
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
        .Add FCONTROL, Asc("T"), conMenu_Edit_AddTemplet
        .Add FCONTROL, Asc("E"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("D"), conMenu_Edit_Delete
        .Add FCONTROL, Asc("M"), conMenu_Edit_ModifyPlanItem
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

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnPlanDataCol As Boolean '��ǰѡ���Ƿ�Ϊ����������
    Dim blnHavePlan As Boolean '��ǰѡ���Ƿ��а���
    Dim bln��ֹԤԼ As Boolean
    Dim blnEnabled As Boolean
    
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = vsfRegistPlan.Rows > vsfRegistPlan.FixedRows
    Case conMenu_EditPopup
        Control.Visible = ((mfrmMain.mFunListActived And (HavePrivs(mstrPrivs, "ģ�����;���ﰲ��")) _
                        Or (mfrmMain.mFunListActived = False And HavePrivs(mstrPrivs, "ģ�����"))))
        Control.Enabled = Control.Visible
    Case conMenu_Edit_AddTemplet '����ģ��
        Control.Visible = HavePrivs(mstrPrivs, "ģ�����") And mfrmMain.mFunListActived
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify '�޸�ģ��
        Control.Visible = HavePrivs(mstrPrivs, "ģ�����") And mfrmMain.mFunListActived
        Control.Enabled = Control.Visible And mlng����ID <> 0
    Case conMenu_Edit_Delete 'ɾ��ģ��
        Control.Visible = HavePrivs(mstrPrivs, "ģ�����") And mfrmMain.mFunListActived
        Control.Enabled = Control.Visible And mlng����ID <> 0

    Case conMenu_Edit_ModifyPlanItem '��������
        Control.Visible = HavePrivs(mstrPrivs, "ģ�����") And mfrmMain.mFunListActived = False
        blnPlanDataCol = vsfRegistPlan.Col >= gPlanGrid_FixedCols
        blnEnabled = mlng����ID > 0
        If zlStr.IsHavePrivs(mstrPrivs, "���п���") = False Then
            'û�С����п��ҡ�Ȩ��ʱ��ֻ�ܵ����������ٴ������Űࡱ�ĺ�Դ
            If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_�Ƿ��ٴ��Ű�)) = "" Then blnEnabled = False
        End If
        Control.Enabled = Control.Visible And blnEnabled And blnPlanDataCol
    Case conMenu_Edit_ModifyUnitRegist '����ԤԼ�Һſ���
        Control.Visible = HavePrivs(mstrPrivs, "ģ�����") And mfrmMain.mFunListActived = False
        blnPlanDataCol = vsfRegistPlan.Col >= gPlanGrid_FixedCols
        blnHavePlan = Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, vsfRegistPlan.Col)) <> ""
        bln��ֹԤԼ = Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col) + 2)) = "-"
        blnEnabled = mlng����ID > 0
        If zlStr.IsHavePrivs(mstrPrivs, "���п���") = False Then
            'û�С����п��ҡ�Ȩ��ʱ��ֻ�ܵ����������ٴ������Űࡱ�ĺ�Դ
            If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_�Ƿ��ٴ��Ű�)) = "" Then blnEnabled = False
        End If
        Control.Enabled = Control.Visible And blnEnabled And blnPlanDataCol And blnHavePlan And Not bln��ֹԤԼ
    Case conMenu_Edit_AllStartNO 'ȫ��������ſ���
        Control.Visible = HavePrivs(mstrPrivs, "ģ�����") And mfrmMain.mFunListActived = False
        Control.Enabled = Control.Visible And mlng����ID <> 0
    Case conMenu_Edit_AllStopNO 'ȫ��ȡ����ſ���
        Control.Visible = HavePrivs(mstrPrivs, "ģ�����") And mfrmMain.mFunListActived = False
        Control.Enabled = Control.Visible And mlng����ID <> 0
    Case conMenu_Edit_ClearAllPlan '�����ǰ��Դ���а���
        Control.Visible = HavePrivs(mstrPrivs, "ģ�����") And mfrmMain.mFunListActived = False
        blnEnabled = mlng����ID > 0
        If zlStr.IsHavePrivs(mstrPrivs, "���п���") = False Then
            'û�С����п��ҡ�Ȩ��ʱ��ֻ�ܵ����������ٴ������Űࡱ�ĺ�Դ
            If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_�Ƿ��ٴ��Ű�)) = "" Then blnEnabled = False
        End If
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ClearAll '������к�Դ����
        Control.Visible = HavePrivs(mstrPrivs, "ģ�����") And mfrmMain.mFunListActived = False
        Control.Enabled = Control.Visible And mlng����ID <> 0
    Case conMenu_Edit_NextMonthNewPlan '�������°���
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And mbytFun = 1
        Control.Enabled = Control.Visible And mlng����ID <> 0
    Case conMenu_Edit_NextWeekNewPlan '�������ܰ���
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And mbytFun = 2
        Control.Enabled = Control.Visible And mlng����ID <> 0
    Case conMenu_View_FindType '���ҷ�ʽ
        Control.Caption = "��" & Decode(mintFindType, 0, "����", 1, "����", 2, "ҽ��", "����") & "���ˡ�"
    Case conMenu_View_Find
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 9 '���ҷ�ʽ
        Control.Checked = Val(Right(Control.ID, 2)) - 1 = mintFindType
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
    Dim frmEdit As frmClinicPlanEdit
    Dim lng��¼ID As Long, lng��ԴId As Long, lng����ID As Long, str���� As String, strItem As String
    Dim obj�����¼ As �����¼, obj�����Դ As �����Դ
    Dim blnFixedRule As Boolean
    
    Err = 0: On Error GoTo errHandler
    lng��ԴId = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ԴID))
    lng����ID = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_����ID))
    lng��¼ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
    strItem = vsfRegistPlan.Cell(flexcpData, 0, vsfRegistPlan.Col)
    str���� = vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_����)
    
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_Modify '�޸�ģ��
        If ModifyPlanInfo(Me, mstrPrivs, mlngModule, mlng����ID) Then Call mfrmMain.NodeChanged("K0_" & mlng����ID)
    Case conMenu_Edit_Delete 'ɾ��ģ��
        If DeletePlan(mstrPrivs, mlng����ID, sccTitle.Tag) Then Call mfrmMain.NodeChanged("")
    Case conMenu_Edit_ModifyPlanItem '����������
        If lng��ԴId <> 0 Or lng����ID <> 0 Then
            Set frmEdit = New frmClinicPlanEdit
            If strItem = "��������" Then strItem = ""
            If frmEdit.ShowMe(Me, 3, Fun_Update, mlng����ID, lng��ԴId, lng����ID, strItem) Then
                Call RefreshOneData
            End If
        End If
    Case conMenu_Edit_ModifyUnitRegist '����������λ
        If lng��ԴId <> 0 Or lng����ID <> 0 Then
            Set frmEdit = New frmClinicPlanEdit
            Call frmEdit.ShowMe(Me, 3, Fun_UpdateUnit, mlng����ID, lng��ԴId, lng����ID, strItem)
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
    Case conMenu_Edit_ClearAllPlan '�����ǰ��Դ����
        If MsgBox("��ȷ��Ҫ�������Ϊ��" & str���� & "�������а�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        If ZlClearPlanBatch(mlng����ID, lng��ԴId) Then Call RefreshOneData
    Case conMenu_Edit_ClearAll '������к�Դ����
        If MsgBox("��ȷ��Ҫ������к�Դ�İ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        If ZlClearPlanBatch(mlng����ID, 0, IIf(HavePrivs(mstrPrivs, "���п���"), 0, UserInfo.ID)) Then
            Call RefreshData(mbytFun, mlng����ID)
        End If
    Case conMenu_Edit_NextWeekNewPlan '�������ܰ���
        Call NextNewPlanByTemplet(mlng����ID, False)
    Case conMenu_Edit_NextMonthNewPlan '�������°���
        Call NextNewPlanByTemplet(mlng����ID, True)
    Case conMenu_View_Refresh
        Call RefreshData(mbytFun, mlng����ID)
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 3 '���ҷ�ʽ
        mintFindType = Val(Right(Control.ID, 2)) - 1
        mcbsMain.RecalcLayout
        txtFind(1).Text = ""
        If txtFind(1).Visible And txtFind(1).Enabled Then txtFind(1).SetFocus
    End Select
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
    Call LoadPlanDataByRecordset(vsfRegistPlan, gPlanGrid_DataStyle.Data_Templet, mrsPlanRecords, 3)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function AddNewPlanTemplet() As String
    '���ܣ�����ģ��
    Dim obj���ﰲ�� As New ���ﰲ��, frmPlanInfoEdit As New frmClinicPlanInfoEdit
    Dim strSQL As String
    
    Err = 0: On Error GoTo errHandler
    
    '����Ƿ��а��»����Ű����Ч��Դ
    If CheckIsHavePlan(3, IIf(zlStr.IsHavePrivs(mstrPrivs, "���п���"), 0, UserInfo.ID)) = False Then
        MsgBox "ģ��ֻ�԰��»����Ű�ĺ�Դ��Ч������ǰ�ް��»����Ű�ĺ�Դ�����ȵ�����������>�ٴ���Դ��������ӣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    obj���ﰲ��.�Ű෽ʽ = 3 '�Ű෽ʽ��0-�̶��Ű�;1-�����Ű�;2-�����Ű�;3-ģ��
    If frmPlanInfoEdit.ShowMe(Me, mlngModule, 1, obj���ﰲ��) = False Then Exit Function
    
    obj���ﰲ��.����ID = zlDatabase.GetNextId("�ٴ������")
    'Zl_�ٴ������_Add
    '  --˵���������������ʱ������һ���޺�Դ��Ϣ���ٴ����ﰲ��
    strSQL = "Zl_�ٴ������_Add("
    '  ��������_In Number,--1-ģ�壬2-�̶�����, 3-�°��ţ�4-�ܰ���
    strSQL = strSQL & "" & 1 & ","
    '  ����id_In   �ٴ������.Id%Type,
    strSQL = strSQL & "" & obj���ﰲ��.����ID & ","
    '  �������_In �ٴ������.�������%Type,
    strSQL = strSQL & "'" & obj���ﰲ��.������� & "',"
    '  վ��_In     ���ű�.վ��%Type,
    strSQL = strSQL & "'" & gstrNodeNo & "',"
    '  ȫԺ��Դ����վ��_In ���ű�.վ��%Type,
    strSQL = strSQL & "'" & gVisitPlan_ModulePara.str��Դά��վ�� & "',"
   '  ����Ա_In   �ٴ����ﰲ��.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ����ʱ��_In �ٴ����ﰲ��.�Ǽ�ʱ��%Type := Null
    strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ","
    '  ��ʼʱ��_In �ٴ����ﰲ��.��ʼʱ��%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  ��ֹʱ��_In �ٴ����ﰲ��.��ֹʱ��%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  ���_In     �ٴ������.���%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  �·�_In     �ٴ������.�·�%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  ����_In     �ٴ������.����%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  Ӧ�÷�Χ_In �ٴ������.Ӧ�÷�Χ%Type := Null,
    strSQL = strSQL & "" & obj���ﰲ��.Ӧ�÷�Χ & ","
    '  ����id_In   �ٴ������.����id%Type := Null,
    strSQL = strSQL & "" & ZVal(obj���ﰲ��.����ID) & ","
    '  ��ע_In     �ٴ������.��ע%Type := Null,
    strSQL = strSQL & "'" & obj���ﰲ��.��ע & "',"
    '  ��Աid_In   ��Ա��.Id%Type := Null),
    strSQL = strSQL & "" & IIf(zlStr.IsHavePrivs(mstrPrivs, "���п���"), "NULL", UserInfo.ID) & ","
    '  ɾ������_In Number := 0,
    strSQL = strSQL & "" & "0" & ","
    '  ģ������_In �ٴ������.ģ������%Type := Null
    strSQL = strSQL & "" & obj���ﰲ��.ģ������ & ")"
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    '�����ģ��ڵ㣺K0_����ID
    AddNewPlanTemplet = "K0_" & obj���ﰲ��.����ID
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ModifyPlanInfo(frmMain As Form, ByVal strPrivs As String, _
    ByVal lngModule As Long, ByVal lng����ID As Long) As Boolean
    '�޸�ģ����Ϣ
    Dim frmPlanInfoEdit As New frmClinicPlanInfoEdit
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim obj���ﰲ�� As New ���ﰲ��
    
    Err = 0: On Error GoTo errHandler
    strSQL = "Select a.ID, a.�Ű෽ʽ, a.�������, a.Ӧ�÷�Χ, a.����id, a.��ע, a.ģ������, a.������" & vbNewLine & _
            "From �ٴ������ A" & vbNewLine & _
            "Where a.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������Ϣ", lng����ID)
    If rsTemp.EOF Then Exit Function
    If zlStr.IsHavePrivs(strPrivs, "���п���") = False Then
        If Nvl(rsTemp!������) <> UserInfo.���� Then
            MsgBox "��û��Ȩ���޸������ƶ���ģ�壡", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    With obj���ﰲ��
        .����ID = Val(Nvl(rsTemp!ID))
        .�Ű෽ʽ = Val(Nvl(rsTemp!�Ű෽ʽ))
        .������� = Nvl(rsTemp!�������)
        .Ӧ�÷�Χ = Val(Nvl(rsTemp!Ӧ�÷�Χ))
        .����ID = Val(Nvl(rsTemp!����ID))
        .��ע = Nvl(rsTemp!��ע)
        .ģ������ = Val(Nvl(rsTemp!ģ������))
    End With
    If frmPlanInfoEdit.ShowMe(frmMain, mlngModule, 1, obj���ﰲ��, True) = False Then Exit Function
    
    '��������
    'Zl_�ٴ������_Update
    strSQL = "Zl_�ٴ������_Update("
    '  ��������_In Number, --1-ģ�壬2-�̶�����
    strSQL = strSQL & "" & 1 & ","
    '  Id_In       �ٴ������.Id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  �������_In �ٴ������.�������%Type := Null,
    strSQL = strSQL & "'" & obj���ﰲ��.������� & "',"
    '  ��ʼʱ��_In �ٴ����ﰲ��.��ʼʱ��%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  ��ֹʱ��_In �ٴ����ﰲ��.��ֹʱ��%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  Ӧ�÷�Χ_In �ٴ������.Ӧ�÷�Χ%Type := Null,
    strSQL = strSQL & "" & obj���ﰲ��.Ӧ�÷�Χ & ","
    '  ����id_In   �ٴ������.����id%Type := Null,
    strSQL = strSQL & "" & ZVal(obj���ﰲ��.����ID) & ","
    '  ��ע_In     �ٴ������.��ע%Type := Null
    strSQL = strSQL & "'" & obj���ﰲ��.������� & "')"
    zlDatabase.ExecuteProcedure strSQL, frmMain.Caption
    
    ModifyPlanInfo = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub RefreshData(ByVal bytFun As Byte, Optional ByVal lng����ID As Long, Optional ByVal blnClear As Boolean)
    '���ܣ�ˢ�°�����������
    '������
    '   strTitle - ������ʾ
    '   lng����ID - ����ID
    '   bytFun - 1-�°��ţ�2-�ܰ���
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim intӦ�÷�Χ As Integer
    
    Err = 0: On Error GoTo errHandler
    If blnClear Then
        mbytFun = bytFun
        mlng����ID = lng����ID
        
        sccTitle.Caption = "���ﰲ��>����ģ��"
        sccTitle.Tag = ""
        lblPlanInfo.Visible = False
        lblPlanInfo.Caption = "Ӧ�÷�Χ��         ��ע��                             "
        lblPlanInfo.ToolTipText = ""
        Set mrsPlanRecords = Nothing
        
        '�ı�˵�����
        Call ZlUpdatePlanMenu(Me, mcbsMain, bytFun, IIf(HavePrivs(mstrPrivs, "���п���"), 0, UserInfo.ID))
    
        strSQL = "Select a.�������, a.Ӧ�÷�Χ, b.���� As Ӧ�ÿ���, a.��ע" & vbNewLine & _
                " From �ٴ������ A, ���ű� B" & vbNewLine & _
                " Where a.����Id = b.Id(+) And a.Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������Ϣ", lng����ID)
        If Not rsTemp.EOF Then
            sccTitle.Caption = "���ﰲ��>" & Nvl(rsTemp!�������)
            sccTitle.Tag = Nvl(rsTemp!�������)
            lblPlanInfo.Visible = True
            intӦ�÷�Χ = Val(Nvl(rsTemp!Ӧ�÷�Χ)) '0-����;1-��������;2-ȫԺͨ��
            lblPlanInfo.Caption = "Ӧ�÷�Χ��" & _
                Decode(intӦ�÷�Χ, 0, "����", 1, "��������(" & Nvl(rsTemp!Ӧ�ÿ���) & ")", "ȫԺ") & _
                "  ��ע��" & Left(Nvl(rsTemp!��ע), 20)
            lblPlanInfo.ToolTipText = Nvl(rsTemp!��ע)
        End If
        Call InitPlanGrid(vsfRegistPlan, gPlanGrid_DataStyle.Data_Templet)
        Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistPlan, Me.Name, "����")
    End If
    
    If lng����ID <> 0 Then
        '��������
        Screen.MousePointer = vbHourglass
        Set mrsPlanRecords = GetPlanRecords(bytFun = 1, lng����ID)
        Call ExecuteFilter
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub RefreshOneData(Optional ByVal blnReLoadData As Boolean = True)
    'ˢ��ָ���к�Դ����
    Dim lng��ԴId As Long, str�շ���Ŀ As String
    
    Err = 0: On Error GoTo errHandle
    '1.��¼ԭ���ݣ�����ȡ������
    With vsfRegistPlan
        lng��ԴId = Val(.TextMatrix(.Row, COL_��ԴID))
        str�շ���Ŀ = .TextMatrix(.Row, COL_��Ŀ)
    End With
    
    If blnReLoadData Then
        '���±��ؼ�¼��
        Set mrsPlanRecords = GetPlanRecords(mbytFun = 1, mlng����ID)
    End If
    
    '2.���½���
    mrsPlanRecords.Filter = "��ԴID=" & lng��ԴId & " And �շ���Ŀ='" & str�շ���Ŀ & "'"
    Call RefreshOnePlanData(vsfRegistPlan, gPlanGrid_DataStyle.Data_Templet, mrsPlanRecords, , , 3)
    mrsPlanRecords.Filter = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Call mfrmMain.ActiveFormChange(Me)
End Sub

Private Sub Form_Load()
    Dim strFindType As String
    Call GetRegInFor(g˽��ģ��, Me.Name, "FindType", strFindType)
    mintFindType = Val(strFindType)
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    
    lblPlanInfo.Move sccTitle.Width - lblPlanInfo.Width - 100, sccTitle.Top + sccTitle.Height - lblPlanInfo.Height - 50
    With vsfRegistPlan
        .Left = 10
        .Top = sccTitle.Top + sccTitle.Height
        .Width = sccTitle.Width
        .Height = Me.ScaleHeight - .Top - 10
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

Private Sub sccTitle_GotFocus()
    On Error Resume Next
    If vsfRegistPlan.Visible And vsfRegistPlan.Enabled Then vsfRegistPlan.SetFocus
End Sub

Private Sub vsfRegistPlan_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    On Error Resume Next
    Call SetPlanGridRangeColor(vsfRegistPlan, gPlanGrid_DataStyle.Data_Templet, mstrOldSelRangePlan)
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
    Dim lng��ԴId As Long, lng����ID As Long
    Dim frmEdit As New frmClinicPlanEdit
    Dim strCurItem As String, blnUpdate As Boolean
    Dim lngCol As Long, lngRow As Long
    Dim strSort As String
    
    Err = 0: On Error GoTo errHandler
    lngCol = vsfRegistPlan.MouseCol
    lngRow = vsfRegistPlan.MouseRow
    If lngRow = 0 Or lngRow = 1 Then
        '����
        If mrsPlanRecords Is Nothing Then Exit Sub
        If mrsPlanRecords.RecordCount = 0 Then Exit Sub
        strSort = GetPlanSortCircleStr(vsfRegistPlan, gPlanGrid_DataStyle.Data_Templet, lngRow, lngCol)
        If strSort <> "" Then
            mrsPlanRecords.Sort = strSort
            Screen.MousePointer = vbHourglass
            Call LoadPlanDataByRecordset(vsfRegistPlan, gPlanGrid_DataStyle.Data_Templet, mrsPlanRecords, 3, , True)
            Screen.MousePointer = vbDefault
        End If
    Else
        lng��ԴId = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_��ԴID))
        lng����ID = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_����ID))
        lngCol = GetPlanItemNameCol(vsfRegistPlan.Col)
        strCurItem = vsfRegistPlan.Cell(flexcpData, 0, lngCol)
        If lng��ԴId = 0 And lng����ID = 0 Then Exit Sub
        If strCurItem = "" Then Exit Sub
        
        blnUpdate = zlStr.IsHavePrivs(mstrPrivs, "ģ�����")
        If zlStr.IsHavePrivs(mstrPrivs, "���п���") = False Then
            'û�С����п��ҡ�Ȩ��ʱ��ֻ�ܵ����������ٴ������Űࡱ�ĺ�Դ
            If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_�Ƿ��ٴ��Ű�)) = "" Then blnUpdate = False
        End If
        If strCurItem = "��������" Then strCurItem = ""
    
        If frmEdit.ShowMe(Me, 3, IIf(blnUpdate, Fun_Update, Fun_View), mlng����ID, lng��ԴId, lng����ID, strCurItem) Then
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

    '�����˵�
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

Private Sub picImgPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    
    vRect = zlControl.GetControlRect(picImgPlan.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsfRegistPlan, lngLeft, lngTop, picImgPlan.Height)
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistPlan, Me.Name, "����")
End Sub

Private Function NextNewPlanByTemplet(ByVal lngģ��ID As Long, Optional ByVal blnMonth As Boolean) As Boolean
    '����ģ�������°���
    Dim strSQL As String, rsTemp As ADODB.Recordset, lng����ID As Long
    Dim intYear As Integer, intMonth As Integer, intWeek As Integer
    Dim dtStart As Date, dtEnd As Date
    Dim strName As String, strKey As String, blnDeletePlan As Boolean
    Dim cllPlan As Collection, i As Integer
    Dim dtCurrent As Date
    
    Err = 0: On Error GoTo errHandler
    If lngģ��ID = 0 Then Exit Function
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
        
        'zl_�ٴ������_Addbytemplet(
        strSQL = "zl_�ٴ������_Addbytemplet("
        'ģ��id_In   �ٴ������.Id%Type,
        strSQL = strSQL & "" & lngģ��ID & ","
        '��Աid_In   ��Ա��.Id%Type,
        strSQL = strSQL & "" & IIf(zlStr.IsHavePrivs(mstrPrivs, "���п���"), "NULL", UserInfo.ID) & ","
        '����id_In   �ٴ������.Id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '�Ű෽ʽ_In �ٴ������.�Ű෽ʽ%Type,
        strSQL = strSQL & "" & IIf(blnMonth, 1, 2) & ","
        '�������_In �ٴ������.�������%Type,
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
        strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ","
        'վ��_In       ���ű�.վ��%Type,
        strSQL = strSQL & "'" & gstrNodeNo & "',"
        'ȫԺ��Դ����վ��_In ���ű�.վ��%Type,
        strSQL = strSQL & "'" & gVisitPlan_ModulePara.str��Դά��վ�� & "',"
        'ɾ������_In Number:=0
        strSQL = strSQL & "" & IIf(blnDeletePlan, 1, 0) & ")"
    
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Next
    If cllPlan.Count > 1 Then gcnOracle.CommitTrans
    
    'XX�³����ڵ㣺K2_���_�·�
    'XX�ܳ����ڵ㣺K3_���_�·�_����
    Call mfrmMain.NodeChanged(strKey)
    NextNewPlanByTemplet = True
    
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

Private Function GetPlanRecords(ByVal blnMonth As Boolean, ByVal lng����ID As Long) As ADODB.Recordset
    '���ܣ���ȡ���ż�¼
    Dim strSQL As String, strSqlSub As String
    Dim strWhere As String
    Dim str������� As String

    Err = 0: On Error GoTo errHandler
    str������� = IIf(gVisitPlan_ModulePara.byt����ȽϷ�ʽ = 0, "a.����", "Lpad(a.����,5,'0')")
    strSqlSub = "       " & str������� & " As �������, a.Id As ��Դid, a.����, a.����, Nvl(a.�Ƿ񽨲���, 0) As �Ƿ񽨲���,a.ԤԼ����, a.����Ƶ��," & vbNewLine & _
                "       Decode(a.���տ���״̬, 1, '����ԤԼ', 2, '��ֹԤԼ', 3, '�ܽڼ������ÿ���', '���ϰ�') As ���տ���״̬," & vbNewLine & _
                "       Nvl(a.�Ƿ��ٴ��Ű�, 0) As �Ƿ��ٴ��Ű�, Decode(a.�Ű෽ʽ, 1, '�����Ű�', 2, '�����Ű�', '�̶��Ű�') As �Ű෽ʽ," & vbNewLine & _
                "       f.���� As ����, f.���� As ���Ҽ���, Nvl(a.�Ƿ���ջ���, 0) As �Ƿ���ջ���," & vbNewLine
    
    'û��"���п���"Ȩ�޵Ĳ���Աֻ�ܲ����Լ��������ҵĺ�Դ
    If HavePrivs(mstrPrivs, "���п���") = False Then
        strWhere = "      And Exists (Select 1 From ������Ա Where ����id = a.����id And ��Աid = [2])"
    End If
    
    strSQL = "Select m.Id, m.����id, m.��Դid, m.��ĿID, m.ҽ��ID, m.ҽ������, m.�Ű����, m.��ʼʱ��, m.��ֹʱ��, n.������Ŀ, n.�ϰ�ʱ��, n.�޺���, n.��Լ��, n.ԤԼ����" & vbNewLine & _
            " From �ٴ����ﰲ�� M, �ٴ��������� N" & vbNewLine & _
            " Where m.Id = n.����id(+) And m.����id = [1] And Nvl(m.�Ű����, 0) <> 6"
    '���ض������Ű�ĵ�����ѯ
    strSQL = strSQL & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select ID, ����id, ��Դid, ��ĿID, ҽ��ID, ҽ������, �Ű����, ��ʼʱ��, ��ֹʱ��," & vbNewLine & _
            "        f_List2str(Cast(Collect(������Ŀ || '' Order By ������Ŀ) As t_Strlist)) As ������Ŀ, �ϰ�ʱ��, �޺���, ��Լ��, ԤԼ����" & vbNewLine & _
            " From (Select Count(1) Over(Partition By m.��Դid, n.������Ŀ) As ���, m.Id, m.����id, m.��Դid, m.��ĿID, m.ҽ��ID, m.ҽ������, m.�Ű����," & vbNewLine & _
            "               m.��ʼʱ��, m.��ֹʱ��, To_Number(RTrim(n.������Ŀ, '��')) As ������Ŀ, n.�ϰ�ʱ��, n.�޺���, n.��Լ��, n.ԤԼ����" & vbNewLine & _
            "        From �ٴ����ﰲ�� M, �ٴ��������� N" & vbNewLine & _
            "        Where m.Id = n.����id(+) And m.����id = [1] And Nvl(m.�Ű����, 0) = 6)" & vbNewLine & _
            " Group By ID, ����id, ���, ��Դid, ��Ŀid, ҽ��id, ҽ������, �Ű����, ��ʼʱ��, ��ֹʱ��, �ϰ�ʱ��, �޺���, ��Լ��, ԤԼ����"
    
    strSQL = "Select b.����id, b.Id As ����id, " & _
            "       Decode(b.ID,Null,e.����,m.����) As �շ���Ŀ, Decode(b.ID,Null,a.ҽ������,b.ҽ������) As ҽ������," & vbNewLine & _
            "       Decode(b.ID,Null,g.����,n.����) As ҽ������, Decode(b.ID,Null,g.רҵ����ְ��,n.רҵ����ְ��) as ҽ��ְ��," & vbNewLine & _
            "       Decode(b.id,Null,i.��ʶ��,j.��ʶ��) as ��ʶ��," & vbNewLine & _
            "        b.�Ű����, b.��ʼʱ��, b.��ֹʱ��," & strSqlSub & vbNewLine & _
            "        b.Id As ��¼id, b.������Ŀ, b.�ϰ�ʱ��, b.�޺���, b.��Լ��, b.ԤԼ���� As ԤԼ���Ʒ�ʽ" & vbNewLine & _
            " From �ٴ������Դ A, (" & strSQL & ") B, �շ���ĿĿ¼ E, ���ű� F, ��Ա�� G, �շ���ĿĿ¼ M, ��Ա�� N,רҵ����ְ�� I,רҵ����ְ�� J" & vbNewLine & _
            " Where a.Id = b.��Դid(+) And a.����id = f.Id" & vbNewLine & _
            "       And a.��Ŀid = e.Id And a.ҽ��ID= g.ID(+) And b.��Ŀid = m.Id(+) And b.ҽ��ID= n.ID(+)" & vbNewLine & _
            "       And g.רҵ����ְ��=i.����(+) And n.רҵ����ְ��=j.����(+)" & vbNewLine & _
            "       And Nvl(a.�Ű෽ʽ, 0) = [3]" & strWhere & vbNewLine & _
            "       And Nvl(a.�Ƿ�ɾ��, 0) = 0 And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & vbNewLine & _
            "       And (e.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or e.����ʱ�� Is Null)" & vbNewLine & _
            "       And (f.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or f.����ʱ�� Is Null)" & vbNewLine & _
            "       And (g.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or g.����ʱ�� Is Null)" & vbNewLine & _
            "       And Nvl(Nvl(f.վ��,[5]),Nvl([4],'-')) = Nvl([4],'-')" & vbNewLine & _
            " Order By " & str������� & ", b.��ʼʱ��, b.��ֹʱ��, b.������Ŀ, b.�ϰ�ʱ��"
    Set GetPlanRecords = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Ű���Ϣ", lng����ID, UserInfo.ID, IIf(blnMonth, 1, 2), _
        gstrNodeNo, gVisitPlan_ModulePara.str��Դά��վ��)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeletePlan(ByVal strPrivs As String, ByVal lng����ID As Long, _
    strTableName As String) As Boolean
    '���ܣ�ɾ�������
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    If lng����ID = 0 Then Exit Function
    
    If zlStr.IsHavePrivs(strPrivs, "���п���") = False Then
        strSQL = "Select 1 From �ٴ������ A" & vbNewLine & _
                " Where a.�Ű෽ʽ = 3 And a.ID=[1] and a.������=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���Ȩ��", lng����ID, UserInfo.����)
        If rsTemp.EOF Then
            MsgBox "��û��Ȩ��ɾ�������ƶ���ģ�壡", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If MsgBox("��ȷ��Ҫɾ����" & strTableName & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Function
    End If
    
    DeletePlan = ZlDeletePlan(lng����ID)
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
