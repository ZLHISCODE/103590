VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmPersonLoanMgr 
   Caption         =   "��Ա������"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9660
   Icon            =   "frmPersonLoanMgr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   135
      ScaleHeight     =   2520
      ScaleWidth      =   5070
      TabIndex        =   1
      Top             =   1290
      Width           =   5070
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4290
         _Version        =   589884
         _ExtentX        =   7567
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6060
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPersonLoanMgr.frx":08CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11959
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
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   1335
      Top             =   255
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
            Picture         =   "frmPersonLoanMgr.frx":115E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonLoanMgr.frx":14B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   45
      Top             =   105
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmPersonLoanMgr.frx":1806
      Left            =   615
      Top             =   165
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPersonLoanMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar
Private mblnUnload As Boolean, mstrPrivs As String, mstrTitle As String    '���ܱ���
Private mlngMode As Long, mstrKey As String
 
Private Enum mPgIndex
    Pg_������� = 250101
    Pg_������� = 250102
End Enum

Private mfrm������ As frmPersonLoanRequisitionMgr
Private mfrm������� As frmPersonOutPayEdit
Private WithEvents mfrmFilter As frmPersonLoanFileter
Attribute mfrmFilter.VB_VarHelpID = -1

Private Const ID_PANE_SEARCH = 0
Private Const ID_PANE_Page = 1
Private mPanSearch As Pane
Private mobjSubFrm As Collection
Private mfrmActive As Form
Private mArrFilter As Variant

Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-09-09 15:04:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    If mfrmFilter Is Nothing Then Set mfrmFilter = New frmPersonLoanFileter
    
    Call mfrmFilter.Init����
     
    With dkpMan
        .ImageList = imlPaneIcons
        Set mPanSearch = .CreatePane(ID_PANE_SEARCH, 400, 400, DockLeftOf, Nothing)
        mPanSearch.Title = "��������": mPanSearch.Options = PaneNoCloseable
        mPanSearch.MinTrackSize.Width = 220: mPanSearch.MaxTrackSize.Width = 300
         Set objPane = .CreatePane(ID_PANE_Page, 400, 400, DockRightOf, mPanSearch)
        objPane.Title = ""
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picList.hwnd
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
End Function

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
'    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case ID_PANE_SEARCH
        Item.Handle = mfrmFilter.hwnd
    Case ID_PANE_Page
        Item.Handle = picList.hwnd
    End Select
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
    If cbsThis.Count >= 2 Then
        blnShowBar = cbsThis(2).Visible
        bytStyle = cbsThis(2).Controls(1).Style
    End If
    
    'ˢ���Ӵ��ڲ˵�
    Call LockWindowUpdate(Me.hwnd)
    
    Me.Caption = "��Ա������ - " & objItem.Caption
    'ɾ�����ڵĹ������������˵���
    For lngCount = cbsThis.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsThis.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsThis.Count To 2 Step -1
        cbsThis(lngCount).Delete
    Next
    
    '�Ӵ������¼���
    mobjSubFrm(CStr(objItem.Tag)).zlDefCommandBars cbsThis
    
    '�ָ����̶���һЩ�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagStretched
    For lngCount = 2 To cbsThis.Count
        cbsThis(lngCount).ContextMenuPresent = False
        cbsThis(lngCount).ShowTextBelowIcons = False
        cbsThis(lngCount).EnableDocking xtpFlagStretched
        For Each objControl In cbsThis(lngCount).Controls
            objControl.Style = bytStyle
        Next
        cbsThis(lngCount).Visible = blnShowBar
    Next
    '�������RecalcLayout����������
    Call LockWindowUpdate(0)
    Set mfrmActive = mobjSubFrm(CStr(tbPage.Selected.Tag))
End Sub

Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2009-09-09 11:01:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo ErrHand:
    
    Set mobjSubFrm = New Collection
    
    Set mfrm������ = New frmPersonLoanRequisitionMgr
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_�������, "�ҵĽ���¼", mfrm������.hwnd, 0)
    objItem.Tag = mPgIndex.Pg_�������
    mobjSubFrm.Add mfrm������, CStr(objItem.Tag)
    
    
    Set mfrm������� = New frmPersonOutPayEdit
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_�������, "�ҵĽ����¼", mfrm�������.hwnd, 0)
    objItem.Tag = mPgIndex.Pg_�������
    mobjSubFrm.Add mfrm�������, CStr(objItem.Tag)
     
     With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long
    Dim lng�ϼ�id  As Long
    Dim blnSucces As Boolean
    If tbPage.Selected Is Nothing Then Exit Sub
    '------------------------------------
    Select Case Control.ID
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
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hwnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_File_BillPrintSet  '����ӡ����
            Call PrintBillSet
        Case Else
            Call mobjSubFrm(CStr(tbPage.Selected.Tag)).zlExecuteCommandBars(Control)
        End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub


Private Sub PrintBillSet()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ʊ�ݴ�ӡ����
    '����:���˺�
    '����:2015-06-30 16:03:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBill As String
    On Error GoTo errHandle
    strBill = "ZL1_BILL_1502"
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Resize()
'    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
'
'    On Error Resume Next
'    cbsThis.GetClientRect Left, Top, Right, Bottom
'    With picList
'        .Left = Left: .Top = Top
'        .Width = Right - Left
'        .Height = Bottom - Top
'    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean
    If tbPage.Selected Is Nothing Then Exit Sub
    If Me.Visible = False Then Exit Sub
    
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case Else
        Call mobjSubFrm(CStr(tbPage.Selected.Tag)).zlUpdateCommandBars(Control)
    End Select
End Sub
Private Sub Form_Activate()
    If mblnUnload Then Unload Me: Exit Sub
End Sub

Private Sub Form_Initialize()
  Call InitCommonControls
End Sub

Private Sub Form_Load()
     Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, False)
     Call InitPanel
     Call InitPage
    '��ʼ�˵���������
End Sub
  
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    SaveWinState Me, App.ProductName, mstrTitle
    '�ر��Ӵ���
    For i = 1 To mobjSubFrm.Count
        If Not mobjSubFrm(i) Is Nothing Then Unload mobjSubFrm(i)
    Next
    
End Sub
Private Function zlPopuMenus(ByVal blnListView As Boolean) As Boolean
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Err = 0: On Error Resume Next
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Function
    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("�����˵�", xtpBarPopup)
    For Each mcbrControl In mcbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mcbrControl.ID, mcbrControl.Caption)
        cbrPopupItem.BeginGroup = mcbrControl.BeginGroup
    Next
    
    If Me.cbsThis.ActiveMenuBar.Controls(3).Visible Then
        Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(3)
        For Each mcbrControl In mcbrMenuBar.CommandBar.Controls
            
            Select Case mcbrControl.ID
            Case conMenu_View_ShowStoped, conMenu_View_ShowAll, conMenu_View_Refresh
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mcbrControl.ID, mcbrControl.Caption)
                cbrPopupItem.BeginGroup = mcbrControl.BeginGroup
                cbrPopupItem.Checked = mcbrControl.Checked
            End Select
        Next
    End If
    cbrPopupBar.ShowPopup
End Function
'
'Private Sub mfrm�������_zlPopupMenus(ByVal blnListView As Boolean)
'   Call zlPopuMenus(blnListView)
'End Sub
'
'Private Sub mfrm������_zlPopupMenus(ByVal blnListView As Boolean)
'   Call zlPopuMenus(blnListView)
'End Sub

Private Function CheckDepend() As Boolean
    '------------------------------------------------------------------------------
    '����:�������������
    '����:���ݺϷ�,����true�����򷵻�False
    '����:���˺�
    '����:2007/08/14
    '------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    CheckDepend = False
    
    On Error GoTo errHandle
    
    gstrSQL = "" & _
    "   Select  B.ID  " & _
    "   From ��Ա����˵�� A, ��Ա�� B " & _
    "   Where A.��Աid = B.ID And A.��Ա���� In ('����Һ�Ա', '�����շ�Ա', 'Ԥ���տ�Ա', 'סԺ����Ա') and B.ID=[1] " & _
    "   Order By ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��鵱ǰ����Ա�Ƿ�Ϊ��Ӧ������Ա", UserInfo.ID)
    If rsTemp.EOF Then
        ShowMsgbox "�㲻�߱�������Һ�Ա�������շ�Ա��Ԥ���տ�Ա��סԺ����Ա�������ʣ�����ʹ�ø�ģ�飡"
        rsTemp.Close
        Exit Function
    End If
    
    gstrSQL = "Select ����   From ���㷽ʽ Where ���� = 1 And Rownum <= 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ֽ���㷽ʽ", UserInfo.ID)
    If rsTemp.EOF Then
        ShowMsgbox "���㷽ʽ�в�����һ�������ֽ����ʵĽ��㷽ʽ,���ڽ��㷽ʽ����������!"
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    CheckDepend = True
     
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub ShowList(ByVal lngMode As Long, ByVal strTitle As String, ByVal frmMain As Form)
    '------------------------------------------------------------------------------
    '����:�������,��ʾ��ص���Ŀ��������Ϣ
    '����:
    '����:���˺�
    '����:2007/08/14
    '------------------------------------------------------------------------------
    mblnUnload = False: mlngMode = lngMode: mstrTitle = strTitle: mstrPrivs = gstrPrivs
    If Not CheckDepend Then Exit Sub            '���������Բ���
    Me.Caption = strTitle
    RestoreWinState Me, App.ProductName, mstrTitle
    
    '2006-04-25:���˺�,ͳһ���ӱ�������ģ��Ĺ���
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    Me.Show , frmMain
    Me.ZOrder 0
End Sub

Public Sub BHShowList(ByVal lngMode As Long, ByVal strTitle As String, ByVal lngMain As Long)
    '------------------------------------------------------------------------------
    '����:�������,��ʾ��ص���Ŀ��������Ϣ
    '����:
    '����:���˺�
    '����:2007/08/14
    '------------------------------------------------------------------------------
    mblnUnload = False: mlngMode = lngMode: mstrTitle = strTitle: mstrPrivs = gstrPrivs
    If Not CheckDepend Then Exit Sub            '���������Բ���
    Me.Caption = strTitle
    RestoreWinState Me, App.ProductName, mstrTitle
    
    '2006-04-25:���˺�,ͳһ���ӱ�������ģ��Ĺ���
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    zlCommFun.ShowChildWindow Me.hwnd, lngMain
    Me.ZOrder 0
End Sub

Private Sub mfrmFilter_zlRefreshCon(ByVal arrFilter As Variant, ByVal blnRequisition As Boolean)
    Set mArrFilter = arrFilter
    '���������˸ı�
    Select Case Val(tbPage.Selected.Tag)
    Case mPgIndex.Pg_�������
        Call mfrm������.zlReLoadData(arrFilter)
    Case mPgIndex.Pg_�������
        Call mfrm�������.zlReLoadData(arrFilter)
    Case Else
    End Select
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
        mfrmFilter.blnRequisition = Val(Item.Tag) = mPgIndex.Pg_�������
    End If
    If Val(Item.Tag) = mPgIndex.Pg_������� Then
        If mfrmFilter.IsMyRequistionConChange Then
            mfrmFilter.ReActionFilter True
        End If
    Else
        If mfrmFilter.IsMyOutPayConChange Then
            mfrmFilter.ReActionFilter False
        End If
    End If
    Call SubWinDefCommandBar(Item)
End Sub
 

