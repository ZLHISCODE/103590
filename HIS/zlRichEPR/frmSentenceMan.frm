VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmSentenceMan 
   Caption         =   "�����ʾ����"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9675
   Icon            =   "frmSentenceMan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   9675
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picApply 
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   3720
      ScaleHeight     =   600
      ScaleWidth      =   5640
      TabIndex        =   2
      Top             =   510
      Width           =   5640
      Begin VB.Label lblApply 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2) ��Χ��"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   4
         Top             =   345
         Width           =   810
      End
      Begin VB.Label lblApply 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1) ˵����"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   90
         Width           =   810
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6810
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSentenceMan.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14182
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
   Begin MSComctlLib.ImageList imgClass 
      Left            =   2910
      Top             =   6045
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
            Picture         =   "frmSentenceMan.frx":0E1C
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceMan.frx":13B6
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   6135
      Left            =   60
      TabIndex        =   1
      Top             =   525
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   10821
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imgClass"
      Appearance      =   0
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmSentenceMan.frx":1950
      Left            =   945
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmSentenceMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conPane_Class = 201
Const conPane_Apply = 202
Const conPane_Words = 203

'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mfrmWords As frmSentenceList
Attribute mfrmWords.VB_VarHelpID = -1

Private mstrPrivs As String     '��ǰʹ����Ȩ�޴�
Private mlngId As Long          '��ǰ����ID
Private WithEvents mfrmSentenceExport As frmSentenceExport
Attribute mfrmSentenceExport.VB_VarHelpID = -1

Private Function zlRefTree(Optional lngID As Long) As Long
    '���ܣ�ˢ��װ��ָ������Ĳ����ļ��嵥������λ��ָ�����ļ���
    Dim rsTemp As New ADODB.Recordset
    Dim objNode As MSComctlLib.Node
    
    gstrSQL = "Select ID, �ϼ�id, ����, ����, ˵��, ��Χ" & vbNewLine & _
            "From �����ʾ����" & vbNewLine & _
            "Start With �ϼ�id Is Null" & vbNewLine & _
            "Connect By Prior ID = �ϼ�id" & vbNewLine & _
            "Order By Level, ����"
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, !���� & "-" & !����, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !ID, !���� & "-" & !����, "close")
            End If
            objNode.Tag = !˵�� & vbCrLf & !��Χ: objNode.Sorted = True: objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
    End With
    If Me.tvwClass.Nodes.Count > 0 Then
        If lngID <> 0 Then
            Me.tvwClass.Nodes("_" & lngID).Selected = True
        Else
            Me.tvwClass.Nodes(1).Selected = True
        End If
        If Me.tvwClass.SelectedItem.Children > 0 Then Me.tvwClass.SelectedItem.Expanded = True
        Call tvwClass_NodeClick(Me.tvwClass.SelectedItem)
    Else
        Call tvwClass_NodeClick(Nothing)
    End If
    zlRefTree = Me.tvwClass.Nodes.Count
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefTree = Me.tvwClass.Nodes.Count
End Function


'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim lngRetuId As Long, strTemp As String
Dim cbrControl As CommandBarControl
            
    Err = 0: On Error GoTo ErrHand
    '------------------------------------
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Call mfrmWords.zlExecuteControl(Control)
    Case conMenu_File_Exit: Unload Me
    
    Case conMenu_File_ExportToXML
        Call mfrmSentenceExport.ShowMe(True, Me)
    Case conMenu_File_ImportFromXML
        Call mfrmSentenceExport.ShowMe(False, Me)
    Case conMenu_Edit_NewParent
        If Me.tvwClass.SelectedItem Is Nothing Then
            If frmSentenceClass.ShowMe(Me, True, Nothing) Then Call zlRefTree
        Else
            If frmSentenceClass.ShowMe(Me, True, Me.tvwClass.SelectedItem) Then Call zlRefTree(Mid(Me.tvwClass.SelectedItem.Key, 2))
        End If
    Case conMenu_Edit_ModifyParent
        If frmSentenceClass.ShowMe(Me, False, Me.tvwClass.SelectedItem) Then Call zlRefTree(Mid(Me.tvwClass.SelectedItem.Key, 2))
    Case conMenu_Edit_DeleteParent
        strTemp = "���ɾ���ôʾ���༰������Ĵʾ�ʾ����" & vbCrLf & "����" & Me.tvwClass.SelectedItem.Text
        If MsgBox(strTemp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "Zl_�����ʾ����_Edit(3," & Mid(Me.tvwClass.SelectedItem.Key, 2) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        lngRetuId = 0
        If Not (Me.tvwClass.SelectedItem.Next Is Nothing) Then
            lngRetuId = Mid(Me.tvwClass.SelectedItem.Next.Key, 2)
        ElseIf Not (Me.tvwClass.SelectedItem.Previous Is Nothing) Then
            lngRetuId = Mid(Me.tvwClass.SelectedItem.Previous.Key, 2)
        ElseIf Not (Me.tvwClass.SelectedItem.Parent Is Nothing) Then
            lngRetuId = Mid(Me.tvwClass.SelectedItem.Parent.Key, 2)
        End If
        Call zlRefTree(lngRetuId)
    
    Case conMenu_Edit_NewItem, conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_Request
        Call mfrmWords.zlExecuteControl(Control)
    
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
    Case conMenu_View_Refresh
        Call zlRefTree(mlngId)
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    End Select
    Exit Sub
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Call mfrmWords.zlUpdateControl(Control)
    Case conMenu_Edit_NewParent
        Control.Enabled = (InStr(1, mstrPrivs, "��ɾ��") > 0)
    Case conMenu_Edit_ModifyParent, conMenu_Edit_DeleteParent
        Control.Enabled = (InStr(1, mstrPrivs, "��ɾ��") > 0)
        If Control.Enabled Then Control.Enabled = Not (Me.tvwClass.SelectedItem Is Nothing)
    Case conMenu_Edit_NewItem
        Control.Enabled = Not (Me.tvwClass.SelectedItem Is Nothing)
        If Control.Enabled Then Call mfrmWords.zlUpdateControl(Control)
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        Call mfrmWords.zlUpdateControl(Control)
    Case conMenu_Edit_Request
        Call mfrmWords.zlUpdateControl(Control)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Class
        Item.Handle = Me.tvwClass.hwnd
    Case conPane_Apply
        Item.Handle = Me.picApply.hwnd
    Case conPane_Words
        Item.Handle = mfrmWords.hwnd
    End Select
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

    mstrPrivs = gstrPrivs
    Me.picApply.Tag = Val(Me.picApply.Height)
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, False)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
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
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "��������ΪXml��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ImportFromXML, "��������Xml��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewParent, "�·���(&N)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyParent, "�����޸�(&K)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_DeleteParent, "����ɾ��(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "�´ʾ�(&A)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�ʾ��޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "�ʾ�ɾ��(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "��������(&Q)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
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
        .Add FCONTROL, Asc("T"), conMenu_Edit_Request
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewParent, "�·���"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyParent, "�����޸�")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_DeleteParent, "����ɾ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "�´ʾ�"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�ʾ��޸�")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "�ʾ�ɾ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "��������"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type = xtpControlButton Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    '-----------------------------------------------------
    '���ôʾ���ʾͣ������
    Dim panThis As Pane, panSub As Pane
    If mfrmWords Is Nothing Then Set mfrmWords = New frmSentenceList
    
    Set panThis = dkpMan.CreatePane(conPane_Class, 150, 800, DockLeftOf, Nothing)
    panThis.Title = "�ʾ����"
    panThis.Options = PaneNoCaption
    Set panThis = dkpMan.CreatePane(conPane_Apply, 450, 80, DockRightOf, Nothing)
    panThis.Title = "�ʾ�˵��"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set panSub = dkpMan.CreatePane(conPane_Words, 450, 800, DockBottomOf, panThis)
    panSub.Title = "�ʾ��б�"
    panSub.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    '����װ��
    Call zlRefTree
    Set mfrmSentenceExport = New frmSentenceExport
End Sub

Private Sub Form_Resize()
    Dim panBase As Pane
    If Me.WindowState = vbMinimized Then Exit Sub
    Set panBase = Me.dkpMan.FindPane(conPane_Apply)
    panBase.MinTrackSize.SetSize 0, Me.picApply.Tag / Screen.TwipsPerPixelY
    panBase.MaxTrackSize.SetSize Me.Width / Screen.TwipsPerPixelX, Me.picApply.Tag / Screen.TwipsPerPixelY
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmWords: Set mfrmWords = Nothing
    Unload mfrmSentenceExport: Set mfrmSentenceExport = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub
Private Sub mfrmSentenceExport_zlRefParentTree()
   Call zlRefTree(0)
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim lngWords As Long
    If Node Is Nothing Then
        mlngId = 0
    Else
        mlngId = Mid(Node.Key, 2)
    End If
    
    lngWords = mfrmWords.zlRefFromClass(Me, mlngId)
    Me.stbThis.Panels(2).Text = "�÷�����" & lngWords & "���ʾ�ʾ��"
    
    Dim strTemp As String, strApply As String
    If Node Is Nothing Then Exit Sub
    Me.lblApply(0).Caption = "1) ˵����" & Split(Node.Tag, vbCrLf)(0)
    
    strTemp = Split(Node.Tag, vbCrLf)(1) & "0000000"
    strApply = ""
    If Mid(strTemp, 1, 1) = "1" Then strApply = strApply & "�����ﲡ��"
    If Mid(strTemp, 2, 1) = "1" Then strApply = strApply & "��סԺ����"
    If Mid(strTemp, 3, 1) = "1" Then strApply = strApply & "�������¼"
    If Mid(strTemp, 4, 1) = "1" Then strApply = strApply & "��������"
    If Mid(strTemp, 5, 1) = "1" Then strApply = strApply & "������֤���뱨��"
    If Mid(strTemp, 6, 1) = "1" Then strApply = strApply & "��֪���ļ�"
    If Mid(strTemp, 7, 1) = "1" Then strApply = strApply & "�����Ʊ���"
    If Mid(strTemp, 8, 1) = "1" Then strApply = strApply & "����������"
    
    If strApply = "" Then
        Me.lblApply(1).Caption = "2) ��Χ��δ����"
    Else
        Me.lblApply(1).Caption = "2) ��Χ��" & Mid(strApply, 2)
    End If
End Sub
