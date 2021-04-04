VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockEPRMonitor 
   Caption         =   "�����������"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   Icon            =   "frmDockEPRMonitor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   9165
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picTimeLimit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   1125
      ScaleHeight     =   2235
      ScaleWidth      =   4110
      TabIndex        =   2
      Top             =   615
      Width           =   4110
   End
   Begin VB.PictureBox picMonitor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   1155
      ScaleHeight     =   2235
      ScaleWidth      =   4110
      TabIndex        =   1
      Top             =   3180
      Width           =   4110
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5745
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDockEPRMonitor.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13282
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
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   6420
      TabIndex        =   3
      Top             =   1260
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   68288515
      CurrentDate     =   39158
   End
   Begin VB.Image imgX 
      Height          =   135
      Left            =   750
      MousePointer    =   7  'Size N S
      Top             =   2925
      Width           =   5445
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   585
      Top             =   75
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmDockEPRMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'����
Private Const conPane_Timing = 1
Private Const conPane_Content = 2
Private Const conMenu_File_CurPatient = 211         '��ǰ����
Private Const conMenu_File_MyPatient = 8338          '�ҵĲ���
Private Const conMenu_File_CurDeptPatient = 803     '���Ҳ���

'����
Private mstrPatiInfo As String
Private mlngPatiId As Long
Private mlngPageId As Long
Private mintKind As Integer
Private mlngDeptId As Long      '��ǰ��������id����һ���ǵ�ǰ���˿���
Private mintType As Integer
Private mstrActiveControl As String
Private WithEvents mfrmEPRAuditMonitor As frmEPRAuditMonitor
Attribute mfrmEPRAuditMonitor.VB_VarHelpID = -1
Private WithEvents mfrmEPRAuditTime As frmEPRAuditTime
Attribute mfrmEPRAuditTime.VB_VarHelpID = -1
Private mintState As Integer   '����״̬��1-��Ժ���ˣ�0-��Ժ����

'######################################################################################################################

Public Sub zlRefList(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intKind As Integer, ByVal lngDeptId As Long, ByVal intType As Integer, ByVal intState As Integer)
    '******************************************************************************************************************
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '������ lngPatiId-����id
    '       lngPageId-��ҳid
    '       intKind-�������ࣺ
    '           1-���ﲡ��(ʵ�ʰ������Ｒ��֤����֪���ļ�)
    '           2-���ﲡ��(ʵ�ʰ������Ｒ��֤����֪���ļ�)
    '           4-������
    '       intType:1-��ǰ���ˣ�2-�ҵĲ��ˣ�3-���Ʋ���
    '******************************************************************************************************************
    Dim lngBalance As Long        'ʱ���
    Dim rsTemp As New ADODB.Recordset
    Dim lngCount As Long
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
        
    mlngPatiId = lngPatiID
    mlngPageId = lngPageId
    mintKind = intKind
    mlngDeptId = lngDeptId
    mintType = intType
    mintState = intState
    mfrmEPRAuditMonitor.zlClearData
    
    Select Case intKind
    Case 1: Me.Caption = "�����������(����)"
    Case 2: Me.Caption = "�����������(סԺ)"
    Case 4: Me.Caption = "�����������(����)"
    End Select
    
    '---------------------------------------------------
    '��ȡ���˻�����Ϣ
    Err = 0
    On Error GoTo errHand
    If mintType = 1 Then
        If intKind = 1 Then
            gstrSQL = "Select r.�����, r.No, r.����, r.�Ա�, r.����, r.�Ǽ�ʱ�� From ���˹Һż�¼ r Where r.Id =[1] And r.��¼����=1  and r.��¼״̬=1"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPageId)
            With rsTemp
                If .RecordCount <= 0 Then MsgBox "�ò��˲����ڣ����ܴ������ݴ���", vbExclamation, gstrSysName: Exit Sub
                mstrPatiInfo = "�����:" & !����� & "(No:" & !NO & ")    ����:" & !���� & "(" & !�Ա� & ")" & _
                            "  ����:" & Format(!�Ǽ�ʱ��, "yyyy-MM-dd hh:mm")
            End With
        Else
            gstrSQL = "Select b.סԺ��, a.����, a.�Ա�, a.����, b.��Ժ���� As ����, b.��Ժ����" & _
                    " From ������Ϣ a, ������ҳ b" & _
                    " Where a.����id = b.����id And b.����id = [1] And Nvl(b.��ҳid, 0) = [2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngPageId)
            With rsTemp
                If .RecordCount <= 0 Then MsgBox "�ò��˲����ڣ����ܸպñ���ݺϲ��ȣ�", vbExclamation, gstrSysName: Exit Sub
                mstrPatiInfo = "סԺ��:" & !סԺ�� & "(��" & lngPageId & "��סԺ)    ����:" & !���� & "(" & !�Ա� & ")" & _
                            "  ����:" & Format(!��Ժ����, "yyyy-MM-dd hh:mm")
            End With
        End If
        
        stbThis.Panels(2).Text = mstrPatiInfo
    End If
    Call mfrmEPRAuditTime.zlRefreshData(lngPatiID, lngPageId, intKind, lngDeptId, intType, intState, Me.dtpEnd.Value)
        
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '******************************************************************************************************************
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '       strSubhead����ӡ�ĸ�����
    '******************************************************************************************************************
    
    If mstrActiveControl = "���ݼ��" Then
        Call mfrmEPRAuditMonitor.zlPrintData(bytMode)
    Else
        Call mfrmEPRAuditTime.zlPrintData(bytMode, mstrPatiInfo)
    End If
    
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strItemKey As String
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview:  Call zlRptPrint(0)
    Case conMenu_File_Print:    Call zlRptPrint(1)
    Case conMenu_File_Excel:    Call zlRptPrint(3)
    Case conMenu_File_Exit:     Unload Me
    Case conMenu_File_CurPatient: Call zlRefList(mlngPatiId, mlngPageId, mintKind, mlngDeptId, 1, mintState): mintType = 1
    Case conMenu_File_MyPatient: Call zlRefList(mlngPatiId, mlngPageId, mintKind, mlngDeptId, 2, mintState): mintType = 2
    Case conMenu_File_CurDeptPatient: Call zlRefList(mlngPatiId, mlngPageId, mintKind, mlngDeptId, 3, mintState): mintType = 3
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.STYLE = IIf(cbrControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh:  Call zlRefList(mlngPatiId, mlngPageId, mintKind, mlngDeptId, mintType, mintState)
    Case conMenu_View_Jump
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    '
    Dim lngScaleLeft As Long
    Dim lngScaleTop  As Long
    Dim lngScaleRight  As Long
    Dim lngScaleBottom  As Long
    
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
        
    If Me.WindowState = 1 Then Exit Sub
    If imgX.Top > (lngScaleBottom - lngScaleTop) - 1000 Then imgX.Top = (lngScaleBottom - lngScaleTop) - 1000
    
    imgX.Left = lngScaleLeft
    imgX.Width = lngScaleRight - lngScaleLeft
    imgX.Height = 45
    
    On Error Resume Next
    
    picTimeLimit.Move imgX.Left, lngScaleTop, lngScaleRight - lngScaleLeft, imgX.Top - lngScaleTop
    picMonitor.Move imgX.Left, imgX.Top + imgX.Height, imgX.Width, lngScaleBottom - imgX.Height - imgX.Top
    

End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
'        If Me.ActiveControl.Name = Me.vfgTiming.Name Then
'            Control.Enabled = (Me.vfgTiming.Rows > Me.vfgContent.FixedRows)
'        Else
'            Control.Enabled = (Me.vfgContent.Rows > Me.vfgContent.FixedRows)
'        End If
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case 0: Control.Visible = IIf(mintState = 2, True, False)
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Timing
        
        Set mfrmEPRAuditTime = New frmEPRAuditTime
        Call mfrmEPRAuditTime.zlInitData(Me)
        Item.Handle = mfrmEPRAuditTime.hWnd
        
    Case conPane_Content
        Set mfrmEPRAuditMonitor = New frmEPRAuditMonitor
        Call mfrmEPRAuditMonitor.zlInitData(Me)
        Item.Handle = mfrmEPRAuditMonitor.hWnd
    End Select
End Sub

Private Sub Form_Load()
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrCustControl As CommandBarControlCustom
    Dim cbrToolBar As CommandBar
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagStretched)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_CurPatient, "��ǰ����(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_MyPatient, "�ҵĲ���(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_CurDeptPatient, "���Ʋ���(&D)")
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
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
        
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_CurPatient, "��ǰ����")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_MyPatient, "�ҵĲ���")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_CurDeptPatient, "���Ʋ���")
        Set cbrControl = .Add(xtpControlLabel, "0", "��Ժʱ��"): cbrControl.BeginGroup = True
        Set cbrCustControl = .Add(xtpControlCustom, "0", "")
            cbrCustControl.Handle = Me.dtpEnd.hWnd
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
        dtpEnd.MaxDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
        dtpEnd.Value = dtpEnd.MaxDate - 7
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.ID <> 0 Then cbrControl.STYLE = xtpButtonIconAndCaption
    Next
        
    Dim lngStyle As Long
    
    Set mfrmEPRAuditTime = New frmEPRAuditTime
    Set mfrmEPRAuditMonitor = New frmEPRAuditMonitor
    
    Call mfrmEPRAuditTime.zlInitData(Me)
    Call mfrmEPRAuditMonitor.zlInitData(Me)
    
    Load mfrmEPRAuditTime
    lngStyle = GetWindowLong(mfrmEPRAuditTime.hWnd, GWL_STYLE)
    Call SetWindowLong(mfrmEPRAuditTime.hWnd, GWL_STYLE, lngStyle Or WS_CHILD)
    Call SetParent(mfrmEPRAuditTime.hWnd, picTimeLimit.hWnd)
    Call MoveWindow(mfrmEPRAuditTime.hWnd, 0, 0, picTimeLimit.ScaleWidth / Screen.TwipsPerPixelX, picTimeLimit.ScaleHeight / Screen.TwipsPerPixelY, 1)
    mfrmEPRAuditTime.Show
    
    Load mfrmEPRAuditMonitor
    lngStyle = GetWindowLong(mfrmEPRAuditMonitor.hWnd, GWL_STYLE)
    Call SetWindowLong(mfrmEPRAuditMonitor.hWnd, GWL_STYLE, lngStyle Or WS_CHILD)
    Call SetParent(mfrmEPRAuditMonitor.hWnd, picMonitor.hWnd)
    Call MoveWindow(mfrmEPRAuditMonitor.hWnd, 0, 0, picMonitor.ScaleWidth / Screen.TwipsPerPixelX, picMonitor.ScaleHeight / Screen.TwipsPerPixelY, 1)
    mfrmEPRAuditMonitor.Show

        
    '-----------------------------------------------------
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)

    Unload mfrmEPRAuditMonitor
    Unload mfrmEPRAuditTime
    
    Set mfrmEPRAuditMonitor = Nothing
    Set mfrmEPRAuditTime = Nothing
End Sub

Private Sub imgX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    imgX.Top = imgX.Top + Y
    
    If imgX.Top < 1500 Then imgX.Top = 1500
    If Me.Height - imgX.Top - imgX.Height < 1000 Then imgX.Top = Me.Height - imgX.Height - 1000
                
    Call cbsThis_Resize
End Sub

Private Sub mfrmEPRAuditMonitor_GotFocus()
    mstrActiveControl = "���ݼ��"
End Sub

Private Sub mfrmEPRAuditTime_AfterDocumentChanged(ByVal lngEPRKey As Long)
        
    Call mfrmEPRAuditMonitor.zlRefreshData(lngEPRKey)
    
End Sub

Private Sub mfrmEPRAuditTime_GotFocus()
    mstrActiveControl = "ʱ��Ҫ��"
End Sub

Private Sub mfrmEPRAuditTime_SelectVfgRow(ByVal strPatiInfo As String)
    stbThis.Panels(2).Text = strPatiInfo
End Sub

Private Sub picMonitor_Resize()
    If Not (mfrmEPRAuditMonitor Is Nothing) Then
        mfrmEPRAuditMonitor.Width = picMonitor.Width
        mfrmEPRAuditMonitor.Height = picMonitor.Height
    End If
End Sub


Private Sub picTimeLimit_Resize()
    If Not (mfrmEPRAuditTime Is Nothing) Then
        mfrmEPRAuditTime.Width = picTimeLimit.Width
        mfrmEPRAuditTime.Height = picTimeLimit.Height
    End If
End Sub

