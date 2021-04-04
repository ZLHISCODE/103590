VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmServiceCenter 
   Caption         =   "���߷�������"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15135
   Icon            =   "frmServiceCenter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15135
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer tmrAuto 
      Enabled         =   0   'False
      Left            =   510
      Top             =   1140
   End
   Begin VB.PictureBox picGuide 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   10275
      ScaleHeight     =   330
      ScaleWidth      =   3000
      TabIndex        =   1
      Top             =   10635
      Width           =   3000
      Begin VB.PictureBox picRed 
         BackColor       =   &H000000FF&
         Height          =   225
         Left            =   2145
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   5
         Top             =   45
         Width           =   225
      End
      Begin VB.PictureBox picBlue 
         BackColor       =   &H00FF0000&
         Height          =   225
         Left            =   1155
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   4
         Top             =   45
         Width           =   225
      End
      Begin VB.PictureBox picBlack 
         BackColor       =   &H00000000&
         Height          =   225
         Left            =   90
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   2
         Top             =   45
         Width           =   225
      End
      Begin VB.Label lblGuide 
         Caption         =   "δ����     �ѻ���     ��ȡ��"
         Height          =   195
         Left            =   375
         TabIndex        =   3
         Top             =   60
         Width           =   2550
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   10590
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   635
      SimpleText      =   $"frmServiceCenter.frx":058A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmServiceCenter.frx":05D1
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16298
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
      Left            =   720
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmServiceCenter.frx":0E65
      Left            =   1260
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmServiceCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String
Public mlngModul As Long
Private mlngFaceBackColor As Long

Private mfrmMessage As frmServiceMessage
Private mfrmList As frmServiceList

Private mWorkPan As Pane '��ǰ����ҳ
Private Enum PaneIdex
    Pane_Left = 1
    Pane_Main = 2
End Enum
Private Enum ShortItemID
    ID_BaseItem = 2
    ID_PlanItem = 1
End Enum

Public mdatBegin As Date, mdatEnd As Date
Private mblnFirst As Boolean

Private Sub dkpMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    dkpMain.RecalcLayout
End Sub

Public Sub LoadData(ByVal lngID As Long)
    Call mfrmList.LoadHistoryData
    Call mfrmList.LoadData(lngID)
End Sub

Private Sub Form_Activate()
    If mblnFirst Then
        mblnFirst = False
        If mfrmMessage.rptMain.SelectedRows.Count = 0 Then Exit Sub
        If mfrmMessage.rptMain.SelectedRows(0).Record Is Nothing Then Exit Sub
        LoadData (mfrmMessage.rptMain.SelectedRows(0).Record(0).Value)
    End If
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo errHandle
    mstrPrivs = gstrPrivs
    mlngModul = 1115
    
    Call InitPara
    Call InitRegist
    Call DefMainCommandBars
    Call InitPanel '��ʼ��dkpMain
    
    Call CreateShortcutBar
    
    mlngFaceBackColor = cbsThis.GetSpecialColor(XPCOLOR_SPLITTER_FACE)
    Me.BackColor = mlngFaceBackColor
    mblnFirst = True
    RestoreWinState mfrmMessage, Me.Caption
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub CreateShortcutBar()
    Err = 0: On Error GoTo errHandle
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function CheckRegistAppointment(ByVal strNO As String) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    strSQL = "Select 1 From ������ü�¼ Where NO = [1] And ��¼����=4 And ����ID Is Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If Not rsTmp.EOF Then
        CheckRegistAppointment = True
        Exit Function
    Else
        '�շѵ�ԤԼ��¼
        strSQL = "Select 1" & vbNewLine & _
                "From ����Ԥ����¼ A, ������ü�¼ B, ���㷽ʽ C" & vbNewLine & _
                "Where a.����id = b.����id And b.No = [1] And b.��¼���� = 4 And a.��¼���� = 4 And a.���㷽ʽ = c.���� And c.���� Not In (7, 8)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        If rsTmp.EOF Then
            CheckRegistAppointment = True
        Else
            MsgBox "��ǰԤԼ��¼�Ѿ�ʹ�÷������������շ�,�޷��ڻ��߷������Ĵ���ȡ��ԤԼ!"
            CheckRegistAppointment = False
        End If
    End If
    
End Function

Public Sub DeleteRecord(ByVal strNO As String, ByVal str����Ա As String, ByVal dat�Ǽ�ʱ�� As Date, ByVal blnApp As Boolean)
    Dim strSQL As String, str����NO As String
    Dim Datsys As Date
    Dim datTmp As Date

    If strNO = "" Then
        MsgBox "��ǰû�йҺ�ԤԼ����ȡ����", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If blnApp Then
        If Not BillOperCheck(1, str����Ա, dat�Ǽ�ʱ��, "ȡ��ԤԼ") Then Exit Sub
        
        If frmRegistEditNew.CancelApp(Me, strNO, 1111, GetPrivFunc(glngSys, 1111), True) = False Then
            Exit Sub
        End If
    Else
        If Not BillOperCheck(1, str����Ա, dat�Ǽ�ʱ��, "�˺�") Then Exit Sub
        
        If CheckBillExistReplenishData(strNO) Then
            MsgBox "ѡ��ĹҺż�¼������ҽ��������㣬����������˺Ų�����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If CheckRegistAppointment(strNO) = False Then
            Exit Sub
        End If
        
        If CheckExecuted(strNO, False) Then
            MsgBox "�Һŵ�" & strNO & "�Ѿ���ҽ��������¹�ҽ��,�����˺ţ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If CheckPriceHaveFee(strNO, str����NO) Then Exit Sub
        
        If ExistFee(strNO) Then
            MsgBox strNO & "�Һŵ��Ĳ����Ѿ������˷���,�����˷Ѳ����˺�.", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If frmRegistEditNew.CancelBill(Me, strNO, 1111, GetPrivFunc(glngSys, 1111), True) = False Then Exit Sub
    
    End If
    
    If mfrmList.tabMain(0).Selected Then
        strSQL = "Zl_���߷�������_����("
        strSQL = strSQL & mfrmList.vsfList.RowData(mfrmList.vsfList.Row) & ","
        strSQL = strSQL & "Null,'"
        strSQL = strSQL & UserInfo.���� & "','"
        strSQL = strSQL & UserInfo.��� & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    If mfrmList.tabMain(3).Selected Then
'        strSQL = "Zl_���߷�������_����("
'        strSQL = strSQL & mfrmMessage.rptMain.SelectedRows.Row(0).Record.Item(6).Value & ","
'        strSQL = strSQL & "Null,'"
'        strSQL = strSQL & UserInfo.���� & "','"
'        strSQL = strSQL & UserInfo.��� & "')"
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitRegist()
    '��ʼ���Һ�
    Dim strDept As String
    Set gobjRegist = New clsRegist
    gobjRegist.zlInitCommon glngSys, gcnOracle, gstrDBUser
    gobjRegist.zlInitData 3, , strDept
End Sub

Private Sub InitPanel()
    Dim objPane As Pane
    
    Err = 0: On Error GoTo errHandle
    Set objPane = dkpMain.CreatePane(Pane_Left, 160, 120, DockLeftOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.Title = "��������Ϣ"
    objPane.Tag = Pane_Left
    objPane.MaxTrackSize.Width = 600
    Set mfrmMessage = New frmServiceMessage
    objPane.Handle = mfrmMessage.Hwnd
    mfrmMessage.ShowMe Me
    
    Set mfrmList = New frmServiceList
    Call mfrmList.InitFrmMain(Me, mstrPrivs)
    Set objPane = dkpMain.CreatePane(Pane_Main, 700, 400, DockRightOf, objPane)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.Tag = Pane_Main
    objPane.Handle = mfrmList.Hwnd
    
    
    
    With dkpMain
        .PanelPaintManager.ClientFrame = xtpTabFrameBorder
        .SetCommandBars cbsThis
        .VisualTheme = ThemeOffice2003
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = True 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
        .PaintManager.HighlighActiveCaption = False
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function DefMainCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-25 15:29:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrSubControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar
    
    Err = 0: On Error GoTo errHandle
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
    cbsThis.ActiveMenuBar.ModifyStyle &H400000, 0 'ȥ���˵���ǰ׺
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Edit, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_Edit
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Bespeak, "ԤԼ�Һ�(&A)")
        cbrControl.IconId = 216
        If InStr(mstrPrivs, ";ԤԼ�ҺŵǼ�;") = 0 Then cbrControl.Visible = False
        If gbytRegistMode = 0 Then
            cbrControl.Visible = False
        Else
            If gdatRegistTime < zlDatabase.Currentdate Then
                cbrControl.Visible = False
            End If
        End If
        Set cbrControl = .Add(xtpControlButton, 3004, "ȡ��ԤԼ(&C)"): cbrControl.BeginGroup = True
        If InStr(mstrPrivs, ";ͣ����Ϣ����;") = 0 And InStr(mstrPrivs, ";ԤԼ�Ǽ���Ϣ����;") = 0 Then cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, 3839, "����(&H)")
        If InStr(mstrPrivs, ";ͣ����Ϣ����;") = 0 Then cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, 3950, "����(&T)")
        If InStr(mstrPrivs, ";ͣ����Ϣ����;") = 0 Then cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, 3936, "ȫ������(&O)")
        If InStr(mstrPrivs, ";ͣ����Ϣ����;") = 0 Then cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, 2601, "��֪ͨ����(&Z)")
        If InStr(mstrPrivs, ";ͣ����Ϣ����;") = 0 And InStr(mstrPrivs, ";ԤԼ�Ǽ���Ϣ����;") = 0 Then cbrControl.Visible = False
        cbrControl.IconId = 11151
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        Set cbrSubControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False)
        cbrSubControl.Checked = True
        Set cbrSubControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False)
        cbrSubControl.Checked = True
        Set cbrSubControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False)
        cbrSubControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        cbrControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "����(&F)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    '����������
    Set cbrToolBar = cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ModifyStyle &H400000, 0 'ȥ���˵���ǰ׺
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Bespeak, "ԤԼ�Һ�(&A)"): cbrControl.BeginGroup = True
        cbrControl.IconId = 216
        If InStr(mstrPrivs, ";ԤԼ�ҺŵǼ�;") = 0 Then cbrControl.Visible = False
        If gbytRegistMode = 0 Then
            cbrControl.Visible = False
        Else
            If gdatRegistTime > zlDatabase.Currentdate Then cbrControl.Visible = False
        End If
        Set cbrControl = .Add(xtpControlButton, 3004, "ȡ��ԤԼ(&C)")
        If InStr(mstrPrivs, ";ͣ����Ϣ����;") = 0 And InStr(mstrPrivs, ";ԤԼ�Ǽ���Ϣ����;") = 0 Then cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, 3839, "����(&T)")
        If InStr(mstrPrivs, ";ͣ����Ϣ����;") = 0 Then cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, 3950, "����(&O)")
        If InStr(mstrPrivs, ";ͣ����Ϣ����;") = 0 Then cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, 2601, "��֪ͨ����(&Z)")
        cbrControl.IconId = 11151
        If InStr(mstrPrivs, ";ͣ����Ϣ����;") = 0 And InStr(mstrPrivs, ";ԤԼ�Ǽ���Ϣ����;") = 0 Then cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "����(&F)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
'    If Not gobjRegist Is Nothing Then gobjRegist.zlDefCommandBars Me, cbsThis, True
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '�����
    With cbsThis.KeyBindings
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add FCONTROL, vbKeyA, conMenu_Edit_SelAll
        .Add FCONTROL, vbKeyC, conMenu_Edit_ClsAll
    End With
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    DefMainCommandBars = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Form_Resize()
    On Error Resume Next
    picGuide.Left = Me.ScaleWidth - 4490
    picGuide.Top = Me.ScaleHeight - 320
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    'ж�����д���
    SaveWinState mfrmMessage, Me.Caption
    Unload mfrmList
    Unload mfrmMessage
'    For i = Forms.Count - 1 To 1 Step -1
'        Unload Forms(i)
'    Next
End Sub
 
Private Sub cbsThis_SpecialColorChanged()
    Me.BackColor = cbsThis.GetSpecialColor(XPCOLOR_SPLITTER_FACE)
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnable As Boolean, i As Integer
    If Me.Visible = False Then Exit Sub
    
    Err = 0: On Error GoTo errHandle
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    With mfrmList
        Select Case Control.ID
            Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
                If .tabMain.Item(1).Selected Or .tabMain.Item(2).Selected Then
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            Case 3839 '����
                If .tabMain.Item(1).Selected Or .tabMain.Item(2).Selected Then
                    Control.Enabled = False
                Else
                    If .tabMain.Item(0).Selected Then
                        If .vsfList.Cell(flexcpForeColor, .vsfList.Row, 1) <> vbRed And .vsfList.TextMatrix(.vsfList.Row, .vsfList.ColIndex("��ϢID")) <> "" Then
                            Control.Enabled = True
                        Else
                            Control.Enabled = False
                        End If
                    Else
                        Control.Enabled = False
                    End If
                End If
            Case 3004 'ȡ��ԤԼ
                If .tabMain.Item(1).Selected Then
                    Control.Enabled = False
                Else
                    If .tabMain.Item(0).Selected Then
                        If .vsfList.Cell(flexcpForeColor, .vsfList.Row, 1) <> vbRed And .vsfList.TextMatrix(.vsfList.Row, .vsfList.ColIndex("��ϢID")) <> "" Then
                            Control.Enabled = True
                        Else
                            Control.Enabled = False
                        End If
                    End If
                    If .tabMain.Item(2).Selected Then
                        If mfrmList.mfrmAppHistory.lblnote.Visible = True Or InStr(mstrPrivs, ";ԤԼ�Ǽ���Ϣ����;") = 0 Then
                            Control.Enabled = False
                        Else
                            Control.Enabled = True
                        End If
                    End If
                    If .tabMain.Item(3).Selected Then
                        If .vsfApp.Cell(flexcpForeColor, .vsfApp.Row, 1) <> vbRed And .vsfApp.TextMatrix(.vsfApp.Row, .vsfApp.ColIndex("���ݺ�")) <> "" Then
                            Control.Enabled = True
                        Else
                            Control.Enabled = False
                        End If
                    End If
                End If
            Case 2601 '��֪ͨ����
                If .tabMain.Item(3).Selected Or .tabMain.Item(2).Selected Then
                    Control.Enabled = False
                Else
                    If .tabMain.Item(0).Selected Then
                        If .vsfList.TextMatrix(.vsfList.Row, .vsfList.ColIndex("�Ƿ���")) = 0 And .vsfList.TextMatrix(.vsfList.Row, .vsfList.ColIndex("��ϢID")) <> "" Then
                            Control.Enabled = True
                        Else
                            Control.Enabled = False
                        End If
                    Else
                        If mfrmMessage.rptMain.SelectedRows.Row(0).Record.Item(1).ForeColor = vbBlack Or mfrmMessage.rptMain.SelectedRows.Row(0).Record.Item(1).ForeColor = -1 Then
                            Control.Enabled = True
                        Else
                            Control.Enabled = False
                        End If
                    End If
                End If
            Case 3950 '����
                If InStr(.tabMain.Selected.Caption, "����") = 0 Then
                    Control.Enabled = False
                Else
                    If .vsfList.TextMatrix(.vsfList.Row, .vsfList.ColIndex("�Ƿ���")) = 0 And .vsfList.TextMatrix(.vsfList.Row, .vsfList.ColIndex("��ϢID")) <> "" And .vsfList.Cell(flexcpForeColor, .vsfList.Row, 1) <> vbRed Then
                        Control.Enabled = True
                    Else
                        Control.Enabled = False
                    End If
                End If
            Case 3936 'ȫ������
                If InStr(.tabMain.Selected.Caption, "����") = 0 Then
                    Control.Enabled = False
                Else
                    blnEnable = False
                    For i = 2 To .vsfList.Rows - 1
                        If .vsfList.TextMatrix(.vsfList.Row, .vsfList.ColIndex("�Ƿ���")) = 0 And .vsfList.TextMatrix(.vsfList.Row, .vsfList.ColIndex("��ϢID")) <> "" And .vsfList.Cell(flexcpForeColor, .vsfList.Row, 1) <> vbRed Then blnEnable = True
                    Next i
                    Control.Enabled = blnEnable
                End If
            Case conMenu_Manage_Bespeak 'ԤԼ�Һ�
                If .tabMain.Item(1).Selected Then
                    Control.Enabled = InStr(mstrPrivs, ";ԤԼ�Ǽ���Ϣ����;") > 0
                Else
                    If InStr(mstrPrivs, ";ԤԼ�ҺŵǼ�;") = 0 Then
                        Control.Visible = False
                    Else
                        gobjRegist.zlUpdateCommandBars Control
                    End If
                End If
            Case Else
                gobjRegist.zlUpdateCommandBars Control
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub ExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call cbsThis_Execute(Control)
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl, lngID As Long
    Dim strNO As String, strOut As String, i As Integer
    Dim strArray() As String, strSQL As String, rsTemp As ADODB.Recordset
    Dim lngPrevID As Long, blnSave As Boolean
    Dim lngPrevIndex As Long, lngRow As Long
    Err = 0: On Error GoTo errHandle
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_File_Exit: Unload Me
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Parameter
        frmServicePara.zlShowMe Me, 1105
        Call InitPara
    Case conMenu_View_Filter
        frmServiceFilter.Show 1, Me
        If frmServiceFilter.mblnOk Then
            With mfrmMessage
                .mdatBegin = frmServiceFilter.dtpBegin.Value
                mdatBegin = frmServiceFilter.dtpBegin.Value
                .mdatEnd = frmServiceFilter.dtpEnd.Value
                mdatEnd = frmServiceFilter.dtpEnd.Value
                .mblnShowRead = frmServiceFilter.chkShowRead.Value
                .mstr�Ǽ��� = NeedName(frmServiceFilter.cbo�Ǽ���.Text)
                .mstr��Ϣ���� = frmServiceFilter.Get��Ϣ����
                .mblnFilter = True
                Call .LoadMessage(False)
                Unload frmServiceFilter
            End With
        End If
    Case conMenu_View_Refresh
        With mfrmMessage
            lngPrevID = 0
            lngPrevIndex = mfrmList.tabMain.Selected.index
            If .rptMain.SelectedRows.Count <> 0 Then
                If Not .rptMain.SelectedRows.Row(0).Record Is Nothing Then
                    lngPrevID = Val(.rptMain.SelectedRows.Row(0).Record.Item(6).Value)
                End If
            End If
            If .mblnFilter = True Then
                Call .LoadMessage(False)
            Else
                Call .LoadMessage(True)
            End If
            Call mfrmList.LoadHistoryData
            .rptMain.SelectedRows.DeleteAll
            If lngPrevID <> 0 Then
                For i = 0 To .rptMain.Rows.Count - 1
                    If Not .rptMain.Rows(i).Record Is Nothing Then
                        If Val(.rptMain.Rows(i).Record(6).Value) = lngPrevID Then
                            .rptMain.SelectedRows.Add .rptMain.Rows(i)
                        End If
                    End If
                Next i
            End If
            If .rptMain.SelectedRows.Count = 0 Then Exit Sub
            If .rptMain.SelectedRows.Row(0).Record Is Nothing Then Exit Sub
            Call LoadData(.rptMain.SelectedRows.Row(0).Record.Item(6).Value)
            mfrmList.tabMain(lngPrevIndex).Selected = True
        End With
    Case conMenu_View_StatusBar
        Control.Checked = Not Control.Checked
        stbThis.Visible = Control.Checked
        picGuide.Visible = stbThis.Visible
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Button
        Control.Checked = Not Control.Checked
        cbsThis(2).Visible = Control.Checked
        Set objControl = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_View_ToolBar_Text, , True)
        objControl.Enabled = Control.Checked
        Set objControl = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_View_ToolBar_Size, , True)
        objControl.Enabled = Control.Checked
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        Control.Checked = Not Control.Checked
        For Each objControl In cbsThis(2).Controls
            objControl.Style = IIf(Control.Checked, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Control.Checked = Not Control.Checked
        cbsThis.Options.LargeIcons = Control.Checked
        cbsThis.RecalcLayout
    Case conMenu_Help_Help: Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.Hwnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.Hwnd)
    Case conMenu_Help_About: Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case 3004 'ȡ��ԤԼ
        If mfrmList.tabMain(0).Selected Then
            lngRow = mfrmList.vsfList.Row
            strNO = mfrmList.vsfList.TextMatrix(mfrmList.vsfList.Row, mfrmList.vsfList.ColIndex("���ݺ�"))
        End If
        If mfrmList.tabMain(3).Selected Then
            strNO = mfrmList.vsfApp.TextMatrix(mfrmList.vsfApp.Row, mfrmList.vsfApp.ColIndex("���ݺ�"))
        End If
        If mfrmList.tabMain(2).Selected Then
            strNO = mfrmList.mfrmAppHistory.cboNO.Text
        End If
        strNO = Trim(strNO)
        strSQL = "Select ����Ա����,�Ǽ�ʱ��,��¼���� From ���˹Һż�¼ Where ��¼״̬=1 And No=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(strNO))
        If rsTemp.EOF Then
            MsgBox "û���ҵ�ԤԼ��¼,����ȡ��ԤԼ!", vbInformation, gstrSysName
            Exit Sub
        Else
            If Val(Nvl(rsTemp!��¼����)) = 1 Then
                Call DeleteRecord(strNO, Nvl(rsTemp!����Ա����), Nvl(rsTemp!�Ǽ�ʱ��), False)
            Else
                Call DeleteRecord(strNO, Nvl(rsTemp!����Ա����), Nvl(rsTemp!�Ǽ�ʱ��), True)
            End If
        End If
        If mfrmMessage.rptMain.SelectedRows.Count <> 0 Then
            If Not mfrmMessage.rptMain.SelectedRows.Row(0).Record Is Nothing Then
                Call LoadData(mfrmMessage.rptMain.SelectedRows.Row(0).Record.Item(6).Value)
                If lngRow <> 0 Then Call mfrmList.LocateNextRecord(lngRow)
            Else
                Call mfrmList.LoadHistoryData
            End If
        Else
            Call mfrmList.LoadHistoryData
        End If
    Case 3950 '����
        If mfrmList.tabMain(0).Selected Then
            lngRow = mfrmList.vsfList.Row
            strNO = mfrmList.vsfList.TextMatrix(mfrmList.vsfList.Row, mfrmList.vsfList.ColIndex("���ݺ�"))
            lngID = mfrmList.vsfList.TextMatrix(mfrmList.vsfList.Row, mfrmList.vsfList.ColIndex("��ϢID"))
        Else
            strNO = mfrmList.vsfApp.TextMatrix(mfrmList.vsfApp.Row, mfrmList.vsfApp.ColIndex("���ݺ�"))
            lngID = mfrmList.vsfApp.TextMatrix(mfrmList.vsfApp.Row, mfrmList.vsfApp.ColIndex("��ϢID"))
        End If
        If MsgBox("�Ƿ�ԹҺŵ�" & strNO & "ȷ������?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then
            Exit Sub
        End If
        Call AffirmChange(Trim(strNO), lngID)
        If mfrmMessage.rptMain.SelectedRows.Count <> 0 Then
            Call LoadData(mfrmMessage.rptMain.SelectedRows.Row(0).Record.Item(6).Value)
            If lngRow <> 0 Then Call mfrmList.LocateNextRecord(lngRow)
        End If
    Case 3936 'ȫ������
        If MsgBox("�Ƿ�����йҺŵ�ȷ������?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then
            Exit Sub
        End If
        If mfrmList.tabMain(0).Selected Then
            With mfrmList.vsfList
                For i = 2 To .Rows - 1
                    If .TextMatrix(i, .ColIndex("�Ƿ���")) = "" Then
                        strNO = .TextMatrix(i, .ColIndex("���ݺ�"))
                        lngID = .TextMatrix(i, .ColIndex("��ϢID"))
                        Call AffirmChange(Trim(strNO), lngID)
                    End If
                Next i
            End With
        Else
            With mfrmList.vsfApp
                For i = 2 To .Rows - 1
                    If .TextMatrix(i, .ColIndex("�Ƿ���")) = "" Then
                        strNO = .TextMatrix(i, .ColIndex("���ݺ�"))
                        lngID = .TextMatrix(i, .ColIndex("��ϢID"))
                        Call AffirmChange(Trim(strNO), lngID)
                    End If
                Next i
            End With
        End If
        If mfrmMessage.rptMain.SelectedRows.Count <> 0 Then
            Call LoadData(mfrmMessage.rptMain.SelectedRows.Row(0).Record.Item(6).Value)
        End If
    Case 2601 '��֪ͨ����
        Call InformPatient
        If mfrmMessage.rptMain.SelectedRows.Count = 0 Then Exit Sub
        If mfrmMessage.rptMain.SelectedRows.Count <> 0 Then
            Call LoadData(mfrmMessage.rptMain.SelectedRows.Row(0).Record.Item(6).Value)
        End If
    Case 3839 '����
        If mfrmList.tabMain.Item(0).Selected = True Then
            With mfrmList.vsfList
                If .TextMatrix(.Row, 1) = "" Then Exit Sub
                lngRow = .Row
                frmServiceChangeNum.InitValue .TextMatrix(.Row, 1), .TextMatrix(.Row, 3), .TextMatrix(.Row, 4), _
                         .TextMatrix(.Row, 5), .TextMatrix(.Row, 6), .TextMatrix(.Row, 9), mfrmList.lblInfo.Caption, _
                         .Cell(flexcpData, .Row, 1)
                frmServiceChangeNum.mlng��ϢID = Val(.RowData(.Row))
                frmServiceChangeNum.Show 1, Me
            End With
        Else
            With mfrmList.vsfList
                lngRow = .Row
                frmServiceChangeNum.InitWithValue .TextMatrix(.Row, 2), .TextMatrix(.Row, 4), .TextMatrix(.Row, 5), _
                         .TextMatrix(.Row, 6), .TextMatrix(.Row, 1), .TextMatrix(.Row, 10), .TextMatrix(.Row, 7), _
                         .TextMatrix(.Row, 9), .Cell(flexcpData, .Row, 1)
                frmServiceChangeNum.mlng��ϢID = Val(.RowData(.Row))
                frmServiceChangeNum.Show 1, Me
            End With
        End If
        If mfrmMessage.rptMain.SelectedRows.Count <> 0 Then
            Call LoadData(mfrmMessage.rptMain.SelectedRows.Row(0).Record.Item(6).Value)
            Call mfrmList.LocateNextRecord(lngRow)
        End If
    Case conMenu_Manage_Bespeak
        blnSave = False
        If mfrmList.tabMain.Item(1).Selected Then  'ԤԼ
            blnSave = mfrmList.DirectApp
        Else
            gobjRegist.zlExecuteCommandBars Me, Control, strOut
            blnSave = True
        End If
        If mfrmMessage.rptMain.SelectedRows.Count <> 0 And blnSave Then
            If Not mfrmMessage.rptMain.SelectedRows.Row(0).Record Is Nothing Then
                Call LoadData(mfrmMessage.rptMain.SelectedRows.Row(0).Record.Item(6).Value)
            End If
        End If
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call zlCallCustomReprot(Val(Split(Control.Parameter, ",")(0)), Trim(Split(Control.Parameter, ",")(1)))
        End If
        gobjRegist.zlExecuteCommandBars Me, Control, strOut
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub NoData()
    mfrmList.tabMain(3).Visible = True
    mfrmList.tabMain(3).Selected = True
    mfrmList.tabMain(2).Visible = False
    mfrmList.tabMain(1).Visible = False
    mfrmList.tabMain(0).Visible = False
End Sub

Public Sub BatchInform()
    '������֪ͨ
    Dim i As Integer, strSQL As String, cllSQL As Collection
    On Error GoTo errH
    Set cllSQL = New Collection
    With mfrmList.vsfList
        For i = 2 To .Rows - 1
            If .TextMatrix(i, 0) = 0 And Val(.RowData(i)) <> 0 And .Cell(flexcpForeColor, i, 1, i, 1) <> vbRed Then
                strSQL = "Zl_���߷�������_����("
                strSQL = strSQL & .TextMatrix(i, .ColIndex("��ϢID")) & ","
                strSQL = strSQL & "Null,'"
                strSQL = strSQL & UserInfo.���� & "','"
                strSQL = strSQL & UserInfo.��� & "')"
                zlAddArray cllSQL, strSQL
            End If
        Next i
        zlExecuteProcedureArrAy cllSQL, Me.Caption
    End With
    If mfrmMessage.rptMain.SelectedRows.Count <> 0 Then
        If Not mfrmMessage.rptMain.SelectedRows.Row(0).Record Is Nothing Then
            Call LoadData(mfrmMessage.rptMain.SelectedRows.Row(0).Record.Item(6).Value)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub BatchCancel()
    '����ȡ��ԤԼ
    Dim i As Integer, strSQL As String, cllSQL As New Collection
    Dim strNO As String
    On Error GoTo errH
    With mfrmList.vsfList
        For i = 2 To .Rows - 1
            strNO = Trim(.TextMatrix(i, .ColIndex("���ݺ�")))
            If .TextMatrix(i, 0) = 0 And Val(.RowData(i)) <> 0 And .Cell(flexcpForeColor, i, 1, i, 1) <> vbRed Then
                'zl_���˹Һż�¼_Delete
                strSQL = "zl_���˹Һż�¼_����_DELETE("
                '  ���ݺ�_In       ������ü�¼.No%Type,
                strSQL = strSQL & "'" & strNO & "',"
                '  ����Ա���_In   ������ü�¼.����Ա���%Type,
                strSQL = strSQL & "'" & UserInfo.��� & "',"
                '  ����Ա����_In   ������ü�¼.����Ա����%Type,
                strSQL = strSQL & "'" & UserInfo.���� & "',"
                '  ժҪ_In         ������ü�¼.ժҪ%Type := Null, --ԤԼȡ��ʱ ��д ���ԤԼȡ��ԭ��
                strSQL = strSQL & "" & "Null)"
                zlAddArray cllSQL, strSQL
                
                strSQL = "Zl_���߷�������_����("
                strSQL = strSQL & .RowData(i) & ","
                strSQL = strSQL & "Null,'"
                strSQL = strSQL & UserInfo.���� & "','"
                strSQL = strSQL & UserInfo.��� & "')"
                zlAddArray cllSQL, strSQL
            End If
        Next i
        zlExecuteProcedureArrAy cllSQL, Me.Caption
    End With
    If mfrmMessage.rptMain.SelectedRows.Count <> 0 Then
        If Not mfrmMessage.rptMain.SelectedRows.Row(0).Record Is Nothing Then
            Call LoadData(mfrmMessage.rptMain.SelectedRows.Row(0).Record.Item(6).Value)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Public Sub InformPatient()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strMsgResult As String, strType As String
    Dim strNO As String
    On Error GoTo errH
    If mfrmList.tabMain(0).Selected Then
        If InStr(mfrmList.tabMain(0).Caption, "ͣ��") > 0 Then
            strType = "ͣ��"
        Else
            strType = "����"
        End If
        If mfrmList.vsfList.TextMatrix(mfrmList.vsfList.Row, mfrmList.vsfList.ColIndex("��ϢID")) = "" Then Exit Sub
        strMsgResult = zlCommFun.ShowMsgbox(gstrSysName, "�Ƿ���֪ͨ��" & _
                                            mfrmList.vsfList.TextMatrix(mfrmList.vsfList.Row, mfrmList.vsfList.ColIndex("����")) & _
                                            "��" & mfrmList.vsfList.TextMatrix(mfrmList.vsfList.Row, mfrmList.vsfList.ColIndex("ԭԤԼʱ��")) & _
                                            "��ԤԼ(" & mfrmList.vsfList.TextMatrix(mfrmList.vsfList.Row, mfrmList.vsfList.ColIndex("���ݺ�")) & _
                                            ")�Ѿ�" & strType & "?", "��֪ͨ����,ȡ��ԤԼ,ȡ��", Me, vbQuestion)
        If strMsgResult = "" Or strMsgResult = "ȡ��" Then Exit Sub
        If strMsgResult = "ȡ��ԤԼ" Then
            strNO = mfrmList.vsfList.TextMatrix(mfrmList.vsfList.Row, mfrmList.vsfList.ColIndex("���ݺ�"))
            strNO = Trim(strNO)
            strSQL = "Select ����Ա����,�Ǽ�ʱ��,��¼���� From ���˹Һż�¼ Where ��¼״̬=1 And No=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(strNO))
            If rsTemp.EOF Then
                MsgBox "û���ҵ�ԤԼ��¼,����ȡ��ԤԼ!", vbInformation, gstrSysName
                Exit Sub
            Else
                If Val(Nvl(rsTemp!��¼����)) = 1 Then
                    Call DeleteRecord(strNO, Nvl(rsTemp!����Ա����), Nvl(rsTemp!�Ǽ�ʱ��), False)
                Else
                    Call DeleteRecord(strNO, Nvl(rsTemp!����Ա����), Nvl(rsTemp!�Ǽ�ʱ��), True)
                End If
            End If
            If mfrmMessage.rptMain.SelectedRows.Count <> 0 Then
                Call LoadData(mfrmMessage.rptMain.SelectedRows.Row(0).Record.Item(6).Value)
            End If
            Exit Sub
        Else
            'ͣ������ҳ��
            strSQL = "Zl_���߷�������_����("
            strSQL = strSQL & mfrmList.vsfList.TextMatrix(mfrmList.vsfList.Row, mfrmList.vsfList.ColIndex("��ϢID")) & ","
            strSQL = strSQL & "Null,'"
            strSQL = strSQL & UserInfo.���� & "','"
            strSQL = strSQL & UserInfo.��� & "')"
        End If
    Else
        If mfrmMessage.rptMain.SelectedRows.Count = 0 Then Exit Sub
        If MsgBox("�Ƿ���֪ͨ��" & mfrmList.mfrmApp.txtName.Text & "��" & mfrmList.mfrmApp.txtTimeBegin.Text & "��" & _
                    mfrmList.mfrmApp.txtTimeEnd.Text & "���ԤԼ�Ǽ��Ѿ�����?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Exit Sub
        'ԤԼ�Ǽ�ҳ��
        strSQL = "Zl_���߷�������_����("
        strSQL = strSQL & mfrmMessage.rptMain.SelectedRows.Row(0).Record.Item(6).Value & ","
        strSQL = strSQL & "Null,'"
        strSQL = strSQL & UserInfo.���� & "','"
        strSQL = strSQL & UserInfo.��� & "')"
    End If
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub AffirmChange(ByVal strNO As String, ByVal lngID As Long)
    On Error GoTo errHandle
    Dim strSQL As String, rsTemp As ADODB.Recordset
    strSQL = "Zl_���߷�������_����("
    strSQL = strSQL & lngID & ",'"
    strSQL = strSQL & strNO & "',"
    strSQL = strSQL & "Null,'"
    strSQL = strSQL & UserInfo.���� & "','"
    strSQL = strSQL & UserInfo.��� & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub zlCallCustomReprot(ByVal lngSys As Long, ByVal strReportNO As String)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ִ�б���
    '���ƣ�������
    '���ڣ�2016-01-11
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Call ReportOpen(gcnOracle, lngSys, strReportNO, Me)
End Sub

Private Sub zlDataPrint(BytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte
    Dim i As Integer
    Dim vsfPrint As VSFlexGrid
    
    Err = 0: On Error GoTo errHandle
    Select Case mfrmList.tabMain.Selected.index
        Case 0
            If InStr(mfrmList.tabMain.Selected.Caption, "ͣ��") > 0 Then
                objOut.Title.Text = "ͣ��ԤԼ�������"
            Else
                objOut.Title.Text = "����ԤԼ�������"
            End If
            Set vsfPrint = mfrmList.vsfList
            vsfPrint.TextMatrix(0, 0) = "�Ƿ�" & vbCrLf & "����"
            vsfPrint.TextMatrix(1, 0) = "�Ƿ�" & vbCrLf & "����"
            vsfPrint.ColWidth(0) = 400
            vsfPrint.ColWidth(vsfPrint.ColIndex("��ϢID")) = 0
            vsfPrint.MergeCells = flexMergeRestrictRows
            For i = 2 To vsfPrint.Rows - 1
                If vsfPrint.TextMatrix(i, 0) = 1 Then
                    vsfPrint.TextMatrix(i, 0) = "��"
                Else
                    vsfPrint.TextMatrix(i, 0) = ""
                End If
                vsfPrint.MergeRow(i) = False
            Next i
            Set objOut.Body = vsfPrint
        Case 3
            objOut.Title.Text = "��ʷԤԼ�嵥"
             Set objOut.Body = mfrmList.vsfApp
        Case Else
            Exit Sub
    End Select
    
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    If BytMode = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, BytMode
    End If
    If mfrmList.tabMain.Selected.index = 0 Then
        vsfPrint.ColWidth(0) = 345
        vsfPrint.TextMatrix(0, 0) = " "
        vsfPrint.TextMatrix(1, 0) = " "
        For i = 2 To vsfPrint.Rows - 1
            If vsfPrint.TextMatrix(i, 0) = "" Then
                vsfPrint.TextMatrix(i, 0) = 0
            Else
                vsfPrint.TextMatrix(i, 0) = 1
            End If
            vsfPrint.MergeRow(i) = False
        Next i
        vsfPrint.MergeCells = flexMergeFixedOnly
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub InitPara()
    Dim strTMP As String
    strTMP = zlDatabase.GetPara("ˢ�·�ʽ", glngSys, 1115, "0", , True) & "|"
    If Split(strTMP, "|")(0) = 0 Then
        tmrAuto.Enabled = False
    Else
        tmrAuto.Enabled = True
        tmrAuto.Interval = Val(Split(strTMP, "|")(1)) * 1000
    End If
End Sub

Private Sub tmrAuto_Timer()
    Dim lngPrevID As Long, lngPrevIndex As Long
    Dim i As Integer, j As Integer
    If Me.Visible = False Then Exit Sub
    With mfrmMessage
        lngPrevID = 0
        lngPrevIndex = mfrmList.tabMain.Selected.index
        If .rptMain.SelectedRows.Count <> 0 Then
            If Not .rptMain.SelectedRows.Row(0).Record Is Nothing Then
                lngPrevID = Val(.rptMain.SelectedRows.Row(0).Record.Item(6).Value)
            End If
        End If
        If .mblnFilter = True Then
            Call .LoadMessage(False)
        Else
            Call .LoadMessage(True)
        End If
        Call mfrmList.LoadHistoryData
        .rptMain.SelectedRows.DeleteAll
        If lngPrevID <> 0 Then
            For i = 0 To .rptMain.Rows.Count - 1
                If Not .rptMain.Rows(i).Record Is Nothing Then
                    If Val(.rptMain.Rows(i).Record(6).Value) = lngPrevID Then
                        .rptMain.SelectedRows.Add .rptMain.Rows(i)
                    End If
                End If
            Next i
        End If
        If .rptMain.SelectedRows.Count = 0 Then Exit Sub
        If .rptMain.SelectedRows.Row(0).Record Is Nothing Then Exit Sub
        Call LoadData(.rptMain.SelectedRows.Row(0).Record.Item(6).Value)
        mfrmList.tabMain(lngPrevIndex).Selected = True
    End With
End Sub
