VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmPreCompendMan 
   Caption         =   "����Ԥ����ٹ���"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8340
   Icon            =   "frmOutlineMan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picCompend 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   225
      ScaleHeight     =   4695
      ScaleWidth      =   4080
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   4080
      Begin MSComctlLib.ImageList imglvw 
         Left            =   3420
         Top             =   3735
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
               Picture         =   "frmOutlineMan.frx":058A
               Key             =   "Custom"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutlineMan.frx":0B24
               Key             =   "Default"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwList 
         Height          =   4140
         Left            =   -15
         TabIndex        =   2
         Top             =   -15
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   7303
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imglvw"
         SmallIcons      =   "imglvw"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtApply 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   4110
         Width           =   3720
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5745
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmOutlineMan.frx":10BE
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9631
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
      Left            =   300
      Top             =   105
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmOutlineMan.frx":1950
      Left            =   960
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPreCompendMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���弶����
Private mstrPrivs As String      '��ǰʹ����Ȩ�޴�
Private WithEvents mfrmDock As frmPhraseMan
Attribute mfrmDock.VB_VarHelpID = -1

'��ʱ����
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim lngCount As Long

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngItemId As Long
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        If Screen.ActiveForm.Name <> mfrmDock.Name Then
            Select Case Control.ID
            Case conMenu_File_Preview: Call zlRptPrint(0)
            Case conMenu_File_Print: Call zlRptPrint(1)
            Case conMenu_File_Excel: Call zlRptPrint(3)
            End Select
        Else
            Call mfrmDock.zlExecuteControl(Control.ID)
        End If
    Case conMenu_File_Exit:     Unload Me
    
    Case conMenu_Edit_NewParent
        lngItemId = 0
        lngItemId = frmPreCompendEdit.ShowMe(Me, True)
        If lngItemId <> 0 Then Call zlRefLists(lngItemId)
    Case conMenu_Edit_NewItem
        Call mfrmDock.zlExecuteControl(Control.ID)
    Case conMenu_Edit_ModifyParent
        If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
        lngItemId = Mid(Me.lvwList.SelectedItem.Key, 2)
        lngItemId = frmPreCompendEdit.ShowMe(Me, False, lngItemId)
        If lngItemId <> 0 Then Call zlRefLists(lngItemId)
    Case conMenu_Edit_Modify
        Call mfrmDock.zlExecuteControl(Control.ID)
    Case conMenu_Edit_DeleteParent
        With Me.lvwList
            If .SelectedItem Is Nothing Then Exit Sub
            If MsgBox("���ɾ����Ԥ�������" & vbCrLf & "����" & .SelectedItem.Text, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSQL = "Zl_����Ԥ�����_Delete(" & Mid(.SelectedItem.Key, 2) & ")"
            Err = 0: On Error GoTo errHand
            Call SQLTest(App.ProductName, Me.Caption, gstrSQL): gcnOracle.Execute gstrSQL, , adCmdStoredProc: Call SQLTest
            Call .ListItems.Remove(.SelectedItem.Key)
            If Not .SelectedItem Is Nothing Then
                Call lvwList_ItemClick(.SelectedItem)
            Else
                Call mfrmDock.zlRefList(0)
            End If
            Me.stbThis.Panels(2).Text = "ʣ��" & .ListItems.Count & "��Ԥ�����"
        End With
        Exit Sub
errHand:
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
        Exit Sub
    Case conMenu_Edit_Delete
        Call mfrmDock.zlExecuteControl(Control.ID)
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Jump
        If Screen.ActiveForm.Name = mfrmDock.Name Then
            Me.lvwList.SetFocus
        Else
            Me.dkpMan.FindPane(2).Select
        End If
    Case conMenu_View_Expend_CurCollapse, conMenu_View_Expend_CurExpend, conMenu_View_Expend_AllCollapse, conMenu_View_Expend_AllExpend
        Call mfrmDock.zlExecuteControl(Control.ID)
    Case conMenu_View_Refresh
        If Me.lvwList.SelectedItem Is Nothing Then
            lngItemId = ""
        Else
            lngItemId = Mid(Me.lvwList.SelectedItem.Key, 2)
        End If
        Call zlRefLists(lngItemId)
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    Dim rptRow As ReportRow
    If Me.Visible = False Then Exit Sub
    
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup
            Control.Visible = Not (InStr(1, mstrPrivs, "�����ɾ��") = 0 And mfrmDock.zlGetPower < 0)
        End Select
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        If Screen.ActiveForm.Name <> mfrmDock.Name Then
            Control.Enabled = Not (Me.lvwList.ListItems.Count = 0)
        Else
            Control.Enabled = (mfrmDock.zlGetRows > 0)
        End If
    Case conMenu_Edit_NewParent
        Control.Visible = Not (InStr(1, mstrPrivs, "�����ɾ��") = 0)
    Case conMenu_Edit_NewItem
        Control.Visible = Not (mfrmDock.zlGetPower < 0)
        Control.Enabled = Not (Me.lvwList.SelectedItem Is Nothing)
        If Control.Enabled Then Control.Enabled = (Val(Mid(Me.lvwList.SelectedItem.Key, 2)) > 0 Or Val(Mid(Me.lvwList.SelectedItem.Key, 2)) = -10)
    Case conMenu_Edit_ModifyParent, conMenu_Edit_DeleteParent
        Control.Visible = Not (InStr(1, mstrPrivs, "�����ɾ��") = 0)
        blnEnabled = Not (Me.lvwList.SelectedItem Is Nothing)
        If blnEnabled Then blnEnabled = (Val(Mid(Me.lvwList.SelectedItem.Key, 2)) > 0)
        Control.Enabled = blnEnabled
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        Control.Visible = Not (mfrmDock.zlGetPower < 0)
        Set rptRow = mfrmDock.zlGetFocusedRow
        blnEnabled = Not (rptRow Is Nothing)
        If blnEnabled Then blnEnabled = Not rptRow.GroupRow
        If blnEnabled Then blnEnabled = (rptRow.Record(0).Value >= mfrmDock.zlGetPower)
        Control.Enabled = blnEnabled
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar:      Control.Checked = Me.stbThis.Visible
    
    Case conMenu_View_Expend_CurCollapse
        Set rptRow = mfrmDock.zlGetFocusedRow
        blnEnabled = Not (rptRow Is Nothing)
        If blnEnabled Then blnEnabled = (rptRow.GroupRow And rptRow.Expanded) Or rptRow.GroupRow = False
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend_CurExpend
        Set rptRow = mfrmDock.zlGetFocusedRow
        blnEnabled = Not (rptRow Is Nothing)
        If blnEnabled Then blnEnabled = (rptRow.GroupRow And rptRow.Expanded = False)
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend_AllCollapse, conMenu_View_Expend_AllExpend
        Control.Enabled = (mfrmDock.zlGetRows > 0)
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = Me.picCompend.hWnd
    Case 2
        If mfrmDock Is Nothing Then
            Set mfrmDock = New frmPhraseMan
            Call mfrmDock.zlSetParent(Me):
            Call mfrmDock.zlShowToolBar(False)
        End If
        Item.Handle = mfrmDock.hWnd
    End Select
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
    mstrPrivs = gstrPrivs
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewParent, "�����(&N)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyParent, "�޸����(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_DeleteParent, "ɾ�����(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "�´ʾ�(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸Ĵʾ�(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ���ʾ�(&L)")
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_Expend, "չ��/�۵���(&X)"): cbrControl.BeginGroup = True
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_Expend_CurCollapse, "�۵���ǰ��(&C)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_Expend_CurExpend, "չ����ǰ��(&E)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_Expend_AllCollapse, "�۵�������(&L)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_Expend_AllExpend, "չ��������(&X)", -1, False
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Jump, "������ת(&J)")
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("N"), conMenu_Edit_NewParent
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F6, conMenu_View_Jump
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Jump
        .AddHiddenCommand conMenu_View_Refresh
    End With
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewParent, "�����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyParent, "�޸����")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_DeleteParent, "ɾ�����")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "�´ʾ�"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸Ĵʾ�")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ���ʾ�")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '����ͣ������
    Dim panCompend As Pane
    Set panCompend = dkpMan.CreatePane(1, 400, 400, DockLeftOf, Nothing)
    panCompend.Title = "Ԥ�����"
    panCompend.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Dim panPhrase As Pane
    Set panPhrase = dkpMan.CreatePane(2, 400, 400, DockRightOf, Nothing)
    panPhrase.Title = "�ʾ�ʾ��"
    panPhrase.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    dkpMan.SetCommandBars Me.cbsThis
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    '����б���̬����
    With Me.lvwList.ColumnHeaders
        .Clear
        .Add , "_����", "����", 2000
        .Add , "_���", "���", 650
        .Add , "_˵��", "˵��", 3000
    End With
    With Me.lvwList
        .ColumnHeaders("_���").Position = 1
        .SortKey = .ColumnHeaders("_���").Index - 1
        .SortOrder = lvwAscending
    End With

    '-----------------------------------------------------
    '����ָ�
    If mfrmDock Is Nothing Then
        Set mfrmDock = New frmPhraseMan
        Call mfrmDock.zlSetParent(Me):
        Call mfrmDock.zlShowToolBar(False)
    End If
    Call RestoreWinState(Me, App.ProductName)

    '-----------------------------------------------------
    '����װ��
    Call zlRefLists
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmDock
    Set mfrmDock = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwList.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwList.SortOrder = IIf(Me.lvwList.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwList.SortKey = ColumnHeader.Index - 1
        Me.lvwList.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwList_DblClick()
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub lvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strApply As String, strText As String
    
    strApply = Mid(Split(Me.lvwList.SelectedItem.Tag, ";")(1), 1, 8): strText = ""
    For lngCount = 1 To Len(strApply)
        If Val(Mid(strApply, lngCount, 1)) = 1 Then
            Select Case lngCount
            Case 1: strText = strText & "�����ﲡ��"
            Case 2: strText = strText & "��סԺ����"
            Case 4: strText = strText & "��������"
            Case 5: strText = strText & "������֤������"
            Case 6: strText = strText & "��֪���ļ�"
            Case 7: strText = strText & "����������"
            Case 8: strText = strText & "�����Ʊ���"
            End Select
        End If
    Next
    If strApply = "" Then
        Me.txtApply.Text = "��δ�����������õĲ������ࡣ"
    Else
        Me.txtApply.Text = "�����������" & Mid(strText, 2) & "��"
    End If
    Me.txtApply.Text = Space(4) & Me.txtApply.Text & IIf(Val(Split(Me.lvwList.SelectedItem.Tag, ";")(0)) = 1, "�ɸ��á�", "")
    If Me.lvwList.Tag = Item.Key Then Exit Sub
    
    If mfrmDock Is Nothing Then Exit Sub
    Me.lvwList.Tag = Item.Key
    lngCount = mfrmDock.zlRefList(Mid(Item.Key, 2))
    Me.stbThis.Panels(2).Text = "����ٹ�" & lngCount & "���ʾ�ʾ��"
End Sub

Private Sub lvwList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call lvwList_DblClick
End Sub

Private Sub lvwList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '�ӱ༭�˵����ƶ��嵯���˵�
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    
    If Button <> vbRightButton Or Me.cbsThis.ActiveMenuBar.Controls(2).Visible <> True Then Exit Sub
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        If cbrControl.Visible And cbrControl.ID <> conMenu_Edit_NewItem Then
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
            cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        End If
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub picCompend_Resize()
    Err = 0: On Error Resume Next
    With Me.lvwList
        .Left = Me.picCompend.ScaleLeft: .Width = Me.picCompend.ScaleWidth - .Left
        .Top = Me.picCompend.ScaleTop: .Height = Me.picCompend.ScaleHeight - Me.txtApply.Height - .Top - 15
    End With
    With Me.txtApply
        .Left = Me.picCompend.ScaleLeft + 15: .Width = Me.picCompend.ScaleWidth - .Left
        .Top = Me.picCompend.ScaleHeight - Me.txtApply.Height
    End With
End Sub

'-------------------------------------------------
'--ͨ�ú������̣�
'-------------------------------------------------

Private Sub zlRefLists(Optional lngItemId As Long)
    '---------------------------------------------
    '��д�б�
    '---------------------------------------------
    Err = 0: On Error GoTo errHand
    
    gstrSQL = "Select Id, �������, �����ı�, ��������, Nvl(�������, 0) As �������, ʹ��ʱ��" & _
            " From �����ļ��ṹ" & _
            " Where �ļ�id Is Null "
    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption)
    Me.lvwList.ListItems.Clear
    With rsTemp
        Do While Not .EOF
            Set objItem = Me.lvwList.ListItems.Add(, "_" & !ID, "" & !�����ı�)
            objItem.SubItems(Me.lvwList.ColumnHeaders("_���").Index - 1) = Format(!�������, "000")
            objItem.SubItems(Me.lvwList.ColumnHeaders("_˵��").Index - 1) = "" & !��������
            objItem.Icon = IIf(!ID < 0, "Default", "Custom"): objItem.SmallIcon = objItem.Icon
            objItem.Tag = !������� & ";" & !ʹ��ʱ��
            If !ID = lngItemId Then objItem.Selected = True
            .MoveNext
        Loop
    End With
    If Me.lvwList.ListItems.Count > 0 Then
        If Me.lvwList.SelectedItem Is Nothing Then Me.lvwList.ListItems(1).Selected = True
        Me.lvwList.SelectedItem.EnsureVisible
        Call lvwList_ItemClick(Me.lvwList.SelectedItem)
        Me.stbThis.Panels(2).Text = "����" & Me.lvwList.ListItems.Count & "��Ԥ�����"
    Else
        Call mfrmDock.zlRefList(0)
        Me.stbThis.Panels(2).Text = ""
    End If
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:��¼���ӡ
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrintLvw
    If Me.lvwList.ListItems.Count = 0 Then Exit Sub
    
    Err = 0: On Error Resume Next
    Set objPrint.Body.objData = Me.lvwList
    objPrint.Title.Text = "Ԥ������嵥"
    objPrint.BelowAppItems.Add "��ӡʱ��:" & Now
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrViewLvw objPrint, bytMode
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub
