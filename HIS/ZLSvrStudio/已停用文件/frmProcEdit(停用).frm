VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmProcEdit 
   Caption         =   "�༭����"
   ClientHeight    =   8016
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   13128
   Icon            =   "frmProcEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8016
   ScaleWidth      =   13128
   StartUpPosition =   1  '����������
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4050
      Index           =   1
      Left            =   1680
      ScaleHeight     =   4056
      ScaleWidth      =   9648
      TabIndex        =   1
      Top             =   3960
      Width           =   9645
      Begin VB.PictureBox picEdit 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   3645
         Left            =   1200
         ScaleHeight     =   3648
         ScaleWidth      =   8616
         TabIndex        =   5
         Top             =   360
         Width           =   8610
         Begin XtremeSyntaxEdit.SyntaxEdit synProcEdit 
            Height          =   2295
            Left            =   360
            TabIndex        =   8
            Top             =   120
            Width           =   2445
            _Version        =   983043
            _ExtentX        =   4313
            _ExtentY        =   4048
            _StockProps     =   84
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            EnableSyntaxColorization=   -1  'True
            ShowLineNumbers =   -1  'True
            ShowSelectionMargin=   -1  'True
            ShowScrollBarVert=   -1  'True
            ShowScrollBarHorz=   -1  'True
            EnableVirtualSpace=   0   'False
            EnableAutoIndent=   -1  'True
            ShowWhiteSpace  =   0   'False
            ShowCollapsibleNodes=   -1  'True
            AutoCompleteWndWidth=   160
         End
         Begin VB.PictureBox picBase 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   3615
            Left            =   3120
            ScaleHeight     =   3588
            ScaleWidth      =   5124
            TabIndex        =   9
            Top             =   0
            Width           =   5145
            Begin VB.TextBox txtLocRow 
               Height          =   300
               Left            =   1125
               TabIndex        =   11
               Top             =   105
               Width           =   1530
            End
            Begin VB.CommandButton cmdProcName 
               Height          =   275
               Left            =   4630
               Picture         =   "frmProcEdit.frx":6852
               Style           =   1  'Graphical
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   898
               Width           =   275
            End
            Begin VB.ComboBox cboProcType 
               Height          =   300
               Left            =   1125
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   495
               Width           =   1530
            End
            Begin VB.ComboBox cboOwner 
               Height          =   300
               Left            =   3525
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   495
               Width           =   1380
            End
            Begin VB.TextBox txtNote 
               Height          =   1260
               Left            =   1125
               MultiLine       =   -1  'True
               TabIndex        =   20
               Top             =   1275
               Width           =   3780
            End
            Begin VB.ComboBox cboProcName 
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   1125
               TabIndex        =   17
               Text            =   "cboProcName"
               Top             =   885
               Width           =   3780
            End
            Begin VB.Label lblNotic 
               AutoSize        =   -1  'True
               Caption         =   "˵�������̱༭��֧�ֿ�ݼ�CTRL+A(ȫѡ)��CTRL+Z(����)��CTRL+C(����)��CTRL+V(ճ��)��CTRL+F(����)��CTRL+H(�滻)��CTRL+G(��λ��)"
               ForeColor       =   &H002222B2&
               Height          =   540
               Left            =   300
               TabIndex        =   22
               Top             =   2760
               Width           =   4815
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblLocRow 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��λ�к�(&G)"
               Height          =   180
               Left            =   120
               TabIndex        =   10
               Top             =   165
               Width           =   990
            End
            Begin VB.Label lblProcType 
               AutoSize        =   -1  'True
               Caption         =   "��������(&T)"
               Height          =   180
               Left            =   120
               TabIndex        =   12
               Top             =   555
               Width           =   990
            End
            Begin VB.Label lblOwner 
               AutoSize        =   -1  'True
               Caption         =   "������(&O)"
               Height          =   180
               Left            =   2685
               TabIndex        =   14
               Top             =   555
               Width           =   810
            End
            Begin VB.Label lblNote 
               AutoSize        =   -1  'True
               Caption         =   "����˵��(&P)"
               Height          =   180
               Left            =   120
               TabIndex        =   19
               Top             =   1335
               Width           =   990
            End
            Begin VB.Label lblProcName 
               AutoSize        =   -1  'True
               Caption         =   "��������(&N)"
               Height          =   180
               Left            =   120
               TabIndex        =   16
               Top             =   945
               Width           =   990
            End
         End
      End
      Begin XtremeSuiteControls.TabControl TbcBase 
         Height          =   975
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   120
         Width           =   1935
         _Version        =   589884
         _ExtentX        =   3413
         _ExtentY        =   1720
         _StockProps     =   64
      End
   End
   Begin VB.Frame fraHSplit 
      Height          =   30
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   21
      Top             =   4800
      Width           =   9615
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2970
      Index           =   0
      Left            =   240
      ScaleHeight     =   2976
      ScaleWidth      =   9696
      TabIndex        =   0
      Top             =   1560
      Width           =   9690
      Begin VB.PictureBox picLast 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   2595
         Left            =   2280
         ScaleHeight     =   2592
         ScaleWidth      =   3576
         TabIndex        =   4
         Top             =   240
         Width           =   3570
         Begin XtremeSyntaxEdit.SyntaxEdit synLastProc 
            Height          =   2295
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   2445
            _Version        =   983043
            _ExtentX        =   4313
            _ExtentY        =   4048
            _StockProps     =   84
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ReadOnly        =   -1  'True
            EnableSyntaxColorization=   -1  'True
            ShowLineNumbers =   -1  'True
            ShowSelectionMargin=   -1  'True
            ShowScrollBarVert=   -1  'True
            ShowScrollBarHorz=   -1  'True
            EnableVirtualSpace=   0   'False
            EnableAutoIndent=   -1  'True
            ShowWhiteSpace  =   0   'False
            ShowCollapsibleNodes=   -1  'True
            AutoCompleteWndWidth=   160
         End
         Begin SHDocVwCtl.WebBrowser wbrCompare 
            Height          =   1935
            Left            =   1080
            TabIndex        =   6
            Top             =   240
            Width           =   1935
            ExtentX         =   3413
            ExtentY         =   3413
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   ""
         End
      End
      Begin XtremeSuiteControls.TabControl TbcBase 
         Height          =   1935
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3015
         _Version        =   589884
         _ExtentX        =   5318
         _ExtentY        =   3413
         _StockProps     =   64
      End
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   7320
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   508
      _ExtentY        =   508
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmProcEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================
'==ģ�����
'==============================================================
Private mobjMain As Object '������
Private mlngKey As Long '�洢����ID
Private mblnOk As Boolean '�Ƿ�ȷ���˳�
Private mptType As ProcType '�洢��������
Private mblnChange As Boolean '�Ƿ������޸�
Private mintState As Integer '�洢����״̬
Private mrsProcedure As ADODB.Recordset '�������嵥
Private mstrProcName As String '�洢��������
Private mblnLoad As Boolean
Private Enum PaneEnum
    PE_��ʷ�䶯 = 0
    PE_��ǰ���� = 1
End Enum
Private mfrmProcedureOwnerCon As frmProcOwnerConn



'==============================================================
'==�����ӿ�
'==============================================================
Public Function ShowMe(ByVal objMain As Object, ByVal lngKey As Long, Optional ByVal ptType As ProcType) As Boolean
'������ptType=�洢���̹��������棬ѡ��Ĺ�������
    Set mobjMain = objMain
    mlngKey = lngKey
    mptType = ptType
    mblnOk = False
    mblnLoad = False
    Me.Show 1, objMain
    ShowMe = mblnOk
End Function

'==============================================================
'==�ؼ��¼�
'==============================================================
Private Sub cboProcName_Click()
    Dim strOwner As String
    If Trim(cboProcName.Text) <> "" And cboProcName.Tag = "" Then
        synProcEdit.Text = gclsBase.GetProgram(Trim(cboProcName.Text), strOwner)
        If strOwner <> "" Then
            cboOwner.Text = strOwner
        End If
    End If
End Sub

Private Sub cboProcName_KeyPress(KeyAscii As Integer)
    If mptType <> ProcType.�û����� And mlngKey = 0 Then
        Call SendMessage(cboProcName.hwnd, CB_SHOWDROPDOWN, 1, 0)
    End If
End Sub

Private Sub cboProcType_Click()
    Select Case cboProcType.ItemData(cboProcType.ListIndex)
        Case ProcType.�䶯����, ProcType.�հ׹���
            If mlngKey = 0 Then
                lblOwner.Visible = False: cboOwner.Visible = False
                LoadProcNames
            Else
                cboOwner.Locked = True: cboProcName.Locked = True
            End If
        Case ProcType.�û�����
            If mlngKey = 0 Then
                lblOwner.Visible = True: cboOwner.Visible = True
                If mptType <> �û����� Then cboProcName.Clear
                cboProcName.Text = "ZLUSER_"
            Else
                cboOwner.Locked = True: cboProcName.Locked = True
            End If
    End Select
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
    Case conMenu_Edit_SaveExit
        If ValidData Then
            If SaveProcData() Then
                mblnOk = True
                Unload Me
            End If
        End If
    Case conMenu_Edit_Save
        If ValidData Then
            If SaveProcData(True) Then
                mblnOk = True
                mblnChange = False
            End If
        End If
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Call Form_Resize
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case conMenu_Edit_Save
            '�û����̲����ݴ�
            Control.Enabled = mblnChange And mptType <> ProcType.�û����� Or mlngKey = 0
        Case conMenu_Edit_SaveExit
            Control.Enabled = mblnChange Or mintState = ProcState.������ Or mlngKey = 0
    End Select
End Sub

Private Sub cmdProcName_Click()
    Dim objText As TextStream
    Dim strSQL As String
    Dim objSQL As clsSQLInfo
    Dim objScript As New clsRunScript
    
    On Error GoTo errH
    cdg.DialogTitle = "ѡ��������ļ�"
    cdg.Filter = "�Զ�������ļ�|*.Sql"
    cdg.flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
    cdg.InitDir = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Path", "Import", App.Path & "\Import\ImportProcdure")
    cdg.FileName = ""
    cdg.MaxFileSize = 32767
    cdg.CancelError = True
    On Error Resume Next
    cdg.ShowOpen
    If err.Number = 0 Then
        On Error GoTo errH
        Me.Refresh
        If cdg.FileTitle <> "" Then
            Set objText = gobjFile.OpenTextFile(cdg.FileName, ForAppending)
            objText.WriteLine "/" '��֤���ڴ洢���̽�����
            objText.Close
            If objScript.OpenFile(cdg.FileName) Then
                Do While Not objScript.EOF
                    If objScript.SQLInfo.Block Then
                        If objScript.SQLInfo.BlockType Like "*PROCEDURE*" Or objScript.SQLInfo.BlockType Like "*FUNCTION*" Then
                            Set objSQL = New clsSQLInfo
                            Call objSQL.CopySQL(objScript.SQLInfo)
                            Exit Do
                        End If
                    End If
                    objScript.ReadNextSQL
                Loop
            Else
                Exit Sub
            End If
            If objSQL Is Nothing Then
                MsgBox "ѡ���ļ���δ������Ч�Ĵ洢���̣�", vbInformation, Me.Caption
                Exit Sub
            End If
            If LoadSQLInfo(objSQL) Then
                mblnChange = True
            End If
        End If
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyG And Shift = vbCtrlMask Then
        Call gclsBase.LocationObj(txtLocRow)
    End If
End Sub

Private Sub Form_Load()
    '�������,���ڴ��建�浼�µĴ���
    fraHSplit.Tag = ""
    wbrCompare.Navigate ("")
    synLastProc.Text = ""
    synProcEdit.Text = ""
    txtNote.Text = "": picPane(PE_��ʷ�䶯).Tag = ""
    cboProcName.Clear: cboProcName.Text = "": cboProcName.Tag = "": cboOwner.Clear
    lblOwner.Visible = True: cboOwner.Visible = True
    cboProcType.Locked = False: cboProcName.Locked = False: cboOwner.Locked = False
    wbrCompare.Visible = True: synLastProc.Visible = True
    fraHSplit.Visible = False: picPane(PE_��ʷ�䶯).Visible = False
    If Not mblnLoad Then
        Call InitCommandBar
    End If
    Call InitTbc
    Call InitSQLArea
    Call FillData
    Call Form_Resize
    mblnChange = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not mblnOk And mblnChange Then
        If MsgBox("�����Ѿ������ı䣬ֱ���˳�����ʧ�޸ġ�ȷ���˳���", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim dbRate As Double, lngTotal As Long
    Dim lngLeft As Long, lngRight As Long, lngTop As Long, lngBottom As Long
    
    On Error Resume Next
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If picPane(PE_��ʷ�䶯).Tag = "����" Then
        picPane(PE_��ǰ����).Move lngLeft, lngTop, lngRight - lngLeft - 30, lngBottom - lngTop - 30
        fraHSplit.Visible = False
        picPane(PE_��ʷ�䶯).Visible = False
    Else
        fraHSplit.Visible = True
        picPane(PE_��ʷ�䶯).Visible = True
        If fraHSplit.Tag = "" Then 'û���϶��Ͱ�Ĭ�ϱ�������
            lngTotal = picPane(PE_��ʷ�䶯).Height + picPane(PE_��ǰ����).Height
            dbRate = picPane(PE_��ʷ�䶯).Height / lngTotal
        Else
            lngTotal = lngBottom - lngTop - 30 - fraHSplit.Height - 30
            dbRate = (fraHSplit.Top) / lngTotal
        End If
        
        If dbRate < 0.1 Then
            dbRate = 0.1
        ElseIf dbRate > 0.9 Then
            dbRate = 0.9
        End If
        lngTotal = lngBottom - lngTop - 30 - fraHSplit.Height - 30
        picPane(PE_��ʷ�䶯).Move lngLeft, lngTop + 30, lngRight - lngLeft, lngTotal * dbRate
        fraHSplit.Move lngLeft, picPane(PE_��ʷ�䶯).Top + picPane(PE_��ʷ�䶯).Height + 15, lngRight - lngLeft, fraHSplit.Height
        picPane(PE_��ǰ����).Move lngLeft, fraHSplit.Top + fraHSplit.Height + 15, lngRight - lngLeft, lngBottom - fraHSplit.Top - fraHSplit.Height - 15
        fraHSplit.Tag = ""
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mrsProcedure Is Nothing) Then Set mrsProcedure = Nothing
    If gobjFile.FolderExists(App.Path & "\Reports") Then Call gobjFile.DeleteFolder(App.Path & "\Reports")
    If Not (mfrmProcedureOwnerCon Is Nothing) Then Unload mfrmProcedureOwnerCon
End Sub

Private Sub fraHSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then fraHSplit.Top = fraHSplit.Top + y
End Sub

Private Sub fraHSplit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraHSplit.Tag = "�϶�"
    Call Form_Resize
End Sub

Private Sub picEdit_Resize()
    On Error Resume Next
    picBase.Move picEdit.ScaleWidth - 30 - picBase.Width, 15, picBase.Width, picEdit.ScaleHeight - 30
    synProcEdit.Move 15, 15, picBase.Left - 15, picEdit.ScaleHeight - 30
End Sub

Private Sub picLast_Resize()
    On Error Resume Next
    wbrCompare.Move 15, 0, picLast.ScaleWidth - 30, picLast.ScaleHeight
    synLastProc.Move 15, 15, picLast.ScaleWidth - 30, picLast.ScaleHeight - 30
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    TbcBase(Index).Move 0, 0, picPane(Index).ScaleWidth, picPane(Index).ScaleHeight
    picPane(Index).Refresh
End Sub

Private Sub synLastProc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        synLastProc.SelectAll 'Ctrl+A
    ElseIf KeyCode = vbKeyC And Shift = vbCtrlMask Then
        synLastProc.Copy
    ElseIf KeyCode = vbKeyF And Shift = vbCtrlMask Then
        synProcEdit.ShowFindReplaceDialog (False)
    End If
End Sub

Private Sub synProcEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        synProcEdit.SelectAll 'Ctrl+A
    ElseIf KeyCode = vbKeyZ And Shift = vbCtrlMask Then
        synProcEdit.UnDo
    ElseIf KeyCode = vbKeyC And Shift = vbCtrlMask Then
        synProcEdit.Copy
    ElseIf KeyCode = vbKeyV And Shift = vbCtrlMask Then
        synProcEdit.Paste
    ElseIf KeyCode = vbKeyF And Shift = vbCtrlMask Then
        synProcEdit.ShowFindReplaceDialog (False)
    ElseIf KeyCode = vbKeyH And Shift = vbCtrlMask Then
        synProcEdit.ShowFindReplaceDialog (True)
    ElseIf KeyCode = vbKeyS And Shift = vbCtrlMask Then
        synProcEdit.ShowFindReplaceDialog (True)
    End If
End Sub

Private Sub txtLocRow_GotFocus()
    Call gclsBase.TxtSelAll(txtLocRow)
End Sub

Private Sub txtLocRow_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    
    If KeyAscii = vbKeyReturn Then
        lngRow = Val(txtLocRow.Text)
        If lngRow = 0 Then lngRow = 1
        If synProcEdit.RowsCount < lngRow Then
            synProcEdit.CurrPos.Row = synProcEdit.RowsCount
        Else
            synProcEdit.CurrPos.Row = lngRow
        End If
        Call gclsBase.LocationObj(txtLocRow)
    Else
        If InStr(1, "0123456789" & Chr(8), Chr(KeyAscii)) <= 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtNote_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub synProcEdit_TextChanged(ByVal nRowFrom As Long, ByVal nRowTo As Long, ByVal nActions As Long)
    mblnChange = True
End Sub
'==============================================================
'==˽�з���
'==============================================================

Private Sub InitCommandBar()
    '******************************************************************************************************************
    '���ܣ���ʼ�˵�������
    '��������
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objExtendedBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
'    cbsMain.DeleteAll
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = frmPubIcons.imgPublic.Icons
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    '------------------------------------------------------------------------------------------------------------------
    '��׼������
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_SaveExit, "���(&S)")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Save, "�ݴ�(&C)")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
    mblnLoad = True
End Sub

Private Sub InitTbc()
    With TbcBase(PE_��ʷ�䶯).PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003
        .BoldSelected = True
        .ClientFrame = xtpTabFrameSingleLine
        .ShowIcons = True
        .DisableLunaColors = False
        .Position = xtpTabPositionTop
        .Appearance = xtpTabAppearanceVisio
        .Color = xtpTabColorOffice2003
        .ColorSet.ButtonSelected = &HFFC0C0     '&HD2BDB6
        .ColorSet.ButtonNormal = &HFFC0C0       '&HD2BDB6
        TbcBase(PE_��ʷ�䶯).RemoveAll
        TbcBase(PE_��ʷ�䶯).InsertItem(0, "�ϴι���", picLast.hwnd, 1).Tag = "�ϴι���"
    End With
    
    With TbcBase(PE_��ǰ����).PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003
        .BoldSelected = True
        .ClientFrame = xtpTabFrameSingleLine
        .ShowIcons = True
        .DisableLunaColors = False
        .Position = xtpTabPositionTop
        .Appearance = xtpTabAppearanceVisio
        .Color = xtpTabColorOffice2003
        .ColorSet.ButtonSelected = &HFFC0C0     '&HD2BDB6
        .ColorSet.ButtonNormal = &HFFC0C0       '&HD2BDB6
        TbcBase(PE_��ǰ����).RemoveAll
        TbcBase(PE_��ǰ����).InsertItem(0, "���ι���", picEdit.hwnd, 1).Tag = "���ι���"
    End With
End Sub

Private Sub InitSQLArea()
    Dim strPath As String, strColor As String
    '�﷨�ؼ���ɫ����
    synLastProc.Font.name = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontName", "Fixedsys")
    synLastProc.Font.Size = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontSize", 12)
    synLastProc.Font.Underline = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontUnderline", 0)
    synLastProc.Font.Italic = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontItalic", 0)
    synLastProc.Font.Bold = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontBold", 0)
    synLastProc.Font.Strikethrough = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontStrikethru", 0)
    synLastProc.BorderStyle = xtpBorderClientEdge
    
    synProcEdit.Font.name = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontName", "Fixedsys")
    synProcEdit.Font.Size = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontSize", 12)
    synProcEdit.Font.Underline = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontUnderline", 0)
    synProcEdit.Font.Italic = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontItalic", 0)
    synProcEdit.Font.Bold = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontBold", 0)
    synProcEdit.Font.Strikethrough = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontStrikethru", 0)
    synProcEdit.BorderStyle = xtpBorderClientEdge
    
    '���ÿؼ�����ʾ��ɫ����Ϊ��SQL
    If Not gblnInIDE Then '���Ӷ໷��֧��
        strPath = App.Path & "\PUBLIC\_sql.schclass"
    Else
        strPath = gobjFile.GetParentFolderName(GetSetting("ZLSOFT", "����ȫ��", "����·��")) & "\PUBLIC\_sql.schclass"
    End If
    If Not gobjFile.FileExists(strPath) Then
        strPath = "C:\Appsoft\PUBLIC\_sql.schclass"
    End If
    If gobjFile.FileExists(strPath) Then
        strColor = ReadFileToString(strPath)
    Else
        strColor = ""
    End If
    synLastProc.SyntaxSet "[Schemes]" & vbCrLf & "SQL" & vbCrLf & "[Themes]" & vbCrLf & "Default" & vbCrLf & "Alternative" & vbCrLf
    synLastProc.SyntaxScheme = strColor
    
    synProcEdit.SyntaxSet "[Schemes]" & vbCrLf & "SQL" & vbCrLf & "[Themes]" & vbCrLf & "Default" & vbCrLf & "Alternative" & vbCrLf
    synProcEdit.SyntaxScheme = strColor
    
End Sub

Private Sub FillData()
    '��ȡ�ֶγ���
    cboProcName.Tag = gclsBase.GetMaxLength("zlProcedure", "����")
    txtNote.MaxLength = gclsBase.GetMaxLength("zlProcedure", "˵��")
    '��Ӵ洢��������
    Call LoadProcType
    '����������
    Call LoadOwner
    '�����ϴι��̻�Ա���Ϣ
    Call LoadProcInfo
    cboProcName.Tag = IIf(synProcEdit.Text <> "", "������", "")
    cboProcType.ListIndex = mptType - 1
    cboProcName.Text = mstrProcName
    If mlngKey <> 0 Then
        cboProcName.AddItem mstrProcName
        cboProcName.ListIndex = 0 '�������ݿ����Դ��
        cboProcType.Locked = True
        cboProcName.Locked = True
    End If
End Sub

Public Function ValidData() As Boolean
    '******************************************************************************************************************
    '���ܣ�У��༭���ݵ���Ч��
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strTMp As String, strCurProc As String
    Dim objCurProc As New clsSQLInfo, objSouce As New clsSQLInfo
    Dim arrTmp As Variant, strSQL As String
    Dim cnOracle As ADODB.Connection
    Dim strPassword As String, strError As String
    Dim rsTmp As ADODB.Recordset, rsCur As ADODB.Recordset
    '����������֤
    If gclsBase.StrIsValid(cboProcName.Text, Val(cboProcName.Tag)) = False Then
        gclsBase.LocationObj cboProcName
        Exit Function
    End If
    If Trim(cboProcName.Text) = "" Then
        MsgBox "�������Ʋ���Ϊ��ֵ���������룡", vbInformation + vbOKOnly, "�������"
        gclsBase.LocationObj cboProcName
        Exit Function
    End If
    '��������֤
    If cboOwner.ListIndex = -1 Then
        MsgBox "��ָ������������!", vbInformation + vbOKOnly, "�������"
        gclsBase.LocationObj cboOwner
        Exit Function
    ElseIf cboOwner.ItemData(cboOwner.ListIndex) = 0 Then
        MsgBox "��ָ�����������ߣ�", vbInformation + vbOKOnly, "�������"
        gclsBase.LocationObj cboOwner
        Exit Function
    End If
    '����������֤
    If cboProcType.ListIndex = -1 Then
        MsgBox "��ָ����������!", vbInformation + vbOKOnly, "�������"
        gclsBase.LocationObj cboProcType
        Exit Function
    End If
    '����˵����֤
    If gclsBase.StrIsValid(txtNote.Text, txtNote.MaxLength) = False Then
        gclsBase.LocationObj txtNote
        Exit Function
    End If
    If mlngKey = 0 Then
        '��֤���������Ƿ�ƥ��
        If Trim(txtNote.Text) = "" Then
            MsgBox "�û����̵Ĺ���˵������Ϊ�գ�", vbInformation + vbOKOnly, "�������"
            gclsBase.LocationObj txtNote
            Exit Function
        End If
    End If
    strTMp = gclsBase.GetProgram(Trim(cboProcName.Text), , True)
    If mptType <> ProcType.�û����� And mlngKey = 0 Then
        If strTMp = "" Then
            MsgBox "�ù��̲���" & IIf(mptType = ProcType.�䶯����, "�䶯����", "�հ׹���") & "��", vbInformation + vbOKOnly, "�������"
            Exit Function
        End If
    End If
    strCurProc = GetCurrentProctext(True)
    If Not objCurProc.LoadSQL(strCurProc, vbCrLf) Or Not objCurProc.Block Then
        MsgBox "�޷������༭����Ĵ洢���̣����ʽ�������±��棡", vbInformation + vbOKOnly, "�������"
        Exit Function
    End If
    Set rsCur = objCurProc.AnsySQL()
    If rsCur Is Nothing Then
        MsgBox "�޷������༭����Ĵ洢���̣����ʽ�������±��棡", vbInformation + vbOKOnly, "�������"
        Exit Function
    End If
    
    If strTMp <> "" Then
        If Not objSouce.LoadSQL(strTMp & vbCrLf & "/", vbCrLf) Then
            MsgBox "�޷��������ݿ��иô洢���̣����ʽ������������ԣ�", vbInformation + vbOKOnly, "�������"
            Exit Function
        End If
        Set rsTmp = objSouce.AnsySQL
        If rsTmp Is Nothing Then
            MsgBox "�޷��������ݿ��иô洢���̣����ʽ������������ԣ�", vbInformation + vbOKOnly, "�������"
            Exit Function
        End If
        If mptType <> ProcType.�û����� Then
            strError = CompareProcPars(rsTmp, rsCur)
            If strError <> "" Then
                MsgBox "�䶯���̻�հ׹��̲�����Ĺ��̲����Լ��������ƻ򷵻�ֵ��������Ϣ���£�" & strError, vbInformation + vbOKOnly, "�������"
                Exit Function
            End If
        End If
    End If
    rsCur.Filter = "λ��=-1"
    If Trim(cboProcName.Text) <> rsCur!���� Then
        MsgBox "�༭����������Ʋ�ƥ�䣡", vbInformation + vbOKOnly, "�������"
        Exit Function
    End If
    '�жϵ�ǰ��¼�û��Ƿ������������ƥ��
    strSQL = "Select User From Dual"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "")
    If Trim(cboOwner.Text) <> rsTmp!User And Not CollectionHave(gcolOwnerConn, "K" & cboOwner.Text) Then
        If mfrmProcedureOwnerCon Is Nothing Then Set mfrmProcedureOwnerCon = New frmProcOwnerConn
        If mfrmProcedureOwnerCon.ShowDialog(Me, cboOwner.Text, strPassword) Then
            
            Set cnOracle = gobjRegister.GetConnection(gstrServer, cboOwner.Text, strPassword, True, OraOLEDB, "", False)
            If cnOracle.State = adStateClosed Then
                Exit Function
            End If
            Call SetSQLTrace(gstrServer, cboOwner.Text, cnOracle)
            gcolOwnerConn.Add cnOracle, "K" & cboOwner.Text
        Else
            Exit Function
        End If
    End If
    ValidData = True
End Function

Private Function SaveProcData(Optional ByVal bln�ݴ� As Boolean) As Boolean
'���ܣ�����洢��������
'������bln�ݴ�-�Ƿ����ݴ�����
    Dim lngKey As Long
    Dim arrSQL() As Variant
    Dim objSQL As New clsSQLInfo
    Dim strTMp As String
    
    On Error GoTo errH
    If mlngKey = 0 Then
        lngKey = gclsBase.GetNextId("zlProcedure")
        If Not bln�ݴ� Then
            strTMp = gclsBase.GetProgram(cboProcName.Text)
        End If
    Else '�޸�
        lngKey = mlngKey
    End If
    Call gclsBase.AddItem(arrSQL, "Zl_Zlprocedure_Update(" & lngKey & "," & mptType & ",'" & cboProcName.Text & "'," & IIf(bln�ݴ�, ProcState.������, ProcState.�ѵ���) & ",'" & txtNote.Text & "','" & cboOwner.Text & "')")
    Call gclsBase.GetProcSQL(lngKey, ProcTextType.�����Զ�����, GetCurrentProctext, arrSQL)
    If strTMp <> "" Then
        Call gclsBase.GetProcSQL(lngKey, ProcTextType.���α�׼����, strTMp, arrSQL)
    End If
    SaveProcData = gclsBase.ExecuteProcedureBeach(gcnOracle, arrSQL, "����洢����")
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
End Function

Private Sub LoadProcInfo()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strProcText As String, pttType As ProcTextType
    Dim i As Long
    
    If mlngKey <> 0 Then
        '��ȡ�洢���̴���
        strSQL = "Select a.Id, a.����, a.����,Upper(a.������) ������, ˵��,״̬,����, ���, b.����" & vbNewLine & _
                        "From Zlprocedure a, Zlproceduretext b" & vbNewLine & _
                        "Where a.Id = b.����id(+) And a.Id = [1]" & vbNewLine & _
                        "Order By b.����, b.���"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", mlngKey)
        If rsTmp.EOF Then
            mlngKey = 0
        Else
            mptType = Val(rsTmp!���� & "")
            mstrProcName = rsTmp!���� & ""
            mintState = Nvl(rsTmp!״̬, 1)
            txtNote.Text = rsTmp!˵�� & ""
            For i = 0 To cboOwner.ListCount
                If cboOwner.List(i) = rsTmp!������ & "" Then
                    cboOwner.ListIndex = i: Exit For
                End If
            Next
        End If
    End If
    If mlngKey <> 0 And mptType <> ProcType.�û����� Then '�û����̲���Ҫ�Ƚϣ�����Ҫ�����ϴι���
        wbrCompare.Visible = True: synLastProc.Visible = False
        '�������Զ����̶�Ӧ���ϴα�׼�����뱾�α�׼���̽��бȽ�
        Call DealWithTmpFolder(True)
        '����ProcTextType.�ϴα�׼����
        rsTmp.Filter = "����=" & ProcTextType.�ϴα�׼����: rsTmp.Sort = "���"
        Call CreateProcText(rsTmp, App.Path & "\Standard\" & mstrProcName & ".sql")
        '����ProcTextType.���α�׼����
        rsTmp.Filter = "����=" & ProcTextType.���α�׼����: rsTmp.Sort = "���"
        Call CreateProcText(rsTmp, App.Path & "\NewStandard\" & mstrProcName & ".sql")
        '����ProcTextType.�����Զ�����
        rsTmp.Filter = "����=" & ProcTextType.�����Զ�����: rsTmp.Sort = "���"
        Call CreateProcText(rsTmp, App.Path & "\ThisProcedure\" & mstrProcName & ".sql")
        '�ϴα�׼�뱾�α�׼�Աȣ����ڶԱ��ļ���������β�ͬ������������Ϊ��ͬ�������߲�ͬʱ����
        If gobjFile.FileExists(App.Path & "\Standard\" & mstrProcName & ".sql") Then
            Call CompareFolder(App.Path & "\Standard", App.Path & "\NewStandard", App.Path & "\Reports")
        End If
        If gobjFile.FileExists(App.Path & "\Reports\" & mstrProcName & ".sql.htm") Then
            TbcBase(PE_��ʷ�䶯).Item(0).Caption = "�ϴα�׼����(��) �� ���α�׼����(��)����Ա�"
            Call wbrCompare.Navigate(App.Path & "\Reports\" & mstrProcName & ".sql.htm")
        Else
            If gobjFile.FileExists(App.Path & "\NewStandard\" & mstrProcName & ".sql") Then
                Call CompareFolder(App.Path & "\NewStandard", App.Path & "\ThisProcedure", App.Path & "\Reports")
            End If
            If gobjFile.FileExists(App.Path & "\Reports\" & mstrProcName & ".sql.htm") Then
                TbcBase(PE_��ʷ�䶯).Item(0).Caption = "���α�׼����(��) �� �����Զ�����(��)����Ա�"
                Call wbrCompare.Navigate(App.Path & "\Reports\" & mstrProcName & ".sql.htm")
            Else
                If gobjFile.FileExists(App.Path & "\Standard\" & mstrProcName & ".sql") Then
                    pttType = ProcTextType.�ϴα�׼����
                ElseIf gobjFile.FileExists(App.Path & "\NewStandard\" & mstrProcName & ".sql") Then
                    pttType = ProcTextType.���α�׼����
                End If
                If pttType <> 0 Then
                    wbrCompare.Visible = False: synLastProc.Visible = True
                    '�����ϴι���
                    rsTmp.Filter = "����=" & pttType
                    rsTmp.Sort = "���"
                    strProcText = ""
                    If Not rsTmp.EOF Then
                        Do While Not rsTmp.EOF
                            strProcText = strProcText & rsTmp!���� & ""
                            rsTmp.MoveNext
                        Loop
                        synLastProc.Text = strProcText
                    End If
                    If strProcText = "" Then picPane(PE_��ʷ�䶯).Tag = "����"
                Else
                    picPane(PE_��ʷ�䶯).Tag = "����"
                End If
            End If
        End If
        '���ɱ��ι���
        rsTmp.Filter = "����=" & ProcTextType.�����Զ�����
        rsTmp.Sort = "���"
        strProcText = ""
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                strProcText = strProcText & rsTmp!���� & ""
                rsTmp.MoveNext
            Loop
            synProcEdit.Text = strProcText
        End If
    Else
        wbrCompare.Visible = False: synLastProc.Visible = False
        picPane(PE_��ʷ�䶯).Tag = "����"
    End If
End Sub

Private Sub CreateProcText(ByRef rsProc As ADODB.Recordset, ByVal strFile As String)
'      blnAdjustName=�Ƿ�ȥ�����̵�����������˫����
    Dim objText As TextStream, strProcText As String
    Dim strName As String
    
    If Not rsProc.EOF Then
        strName = rsProc!���� & ""
        Do While Not rsProc.EOF
            If rsProc!��� = 1 Then
                '���ƴ�˫���ţ���ȥ��
                If UCase(rsProc!����) Like "*" & """" & UCase(strName) & """" & "*" Then
                    strProcText = strProcText & Replace(UCase(rsProc!����), """" & UCase(strName) & """", strName)
                Else
                    strProcText = strProcText & rsProc!����
                End If
            Else
                strProcText = strProcText & rsProc!����
            End If
            rsProc.MoveNext
        Loop
        Set objText = gobjFile.CreateTextFile(strFile)
        objText.Write strProcText
        objText.Close
    End If
End Sub

Private Sub LoadProcNames()
'���ܣ����ش洢����
    Dim strSQL As String
    
    cboProcName.Clear
    If mrsProcedure Is Nothing Then
        strSQL = "Select Object_Name,Owner" & vbNewLine & _
                        "From All_Objects a" & vbNewLine & _
                        "Where a.Owner In (Select Distinct ������ From Zlsystems) And a.Object_Type In ('PROCEDURE', 'FUNCTION') And" & vbNewLine & _
                        "      a.Object_Name Not In (Select Upper(����) ���� From Zlprocedure) And a.Object_Name Not Like 'ZL%_UPGRADECHECK'" & vbNewLine & _
                        "Order By Object_Name"
        Set mrsProcedure = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ�����嵥")
    End If
    mrsProcedure.Filter = "": mrsProcedure.Sort = "Object_Name"
    If mrsProcedure.RecordCount > 0 Then mrsProcedure.MoveFirst
    Do While Not mrsProcedure.EOF
        cboProcName.AddItem mrsProcedure!Object_Name
        mrsProcedure.MoveNext
    Loop
'    If mrsProcedure.RecordCount > 0 Then cboProcName.ListIndex = 0
End Sub

Private Sub LoadOwner()
'���ܣ�����������
    Dim strSQL As String, rsTmp As ADODB.Recordset
    '��ȡ������
    strSQL = "Select Distinct Upper(������)  ������ from zlSystems a"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ������")
    With cboOwner
        If Not rsTmp.EOF Then
            .AddItem "--������--"
            .ItemData(.NewIndex) = 0
            Do While Not rsTmp.EOF
                .AddItem rsTmp!������ & ""
                .ItemData(.NewIndex) = .NewIndex
                rsTmp.MoveNext
            Loop
            .ListIndex = -1
        End If
    End With
End Sub

Private Sub LoadProcType()
'���ܣ����ع�������
    With cboProcType
        .Clear
        .AddItem "1-�䶯����"
        .ItemData(.NewIndex) = 1
        .AddItem "2-�հ׹���"
        .ItemData(.NewIndex) = 2
        .AddItem "3-�û�����"
        .ItemData(.NewIndex) = 3
        .ListIndex = -1
    End With
End Sub

Private Sub DealWithTmpFolder(Optional ByVal blnCreate As Boolean)
'���ܣ�������ʱĿ¼
    'ת��Ϊ��д�Ľű�
    If gobjFile.FolderExists(App.Path & "\Standard") Then Call gobjFile.DeleteFolder(App.Path & "\Standard", True)
    If gobjFile.FolderExists(App.Path & "\NewStandard") Then Call gobjFile.DeleteFolder(App.Path & "\NewStandard", True)
    If gobjFile.FolderExists(App.Path & "\ThisProcedure") Then Call gobjFile.DeleteFolder(App.Path & "\ThisProcedure", True)
    If gobjFile.FolderExists(App.Path & "\Reports") Then Call gobjFile.DeleteFolder(App.Path & "\Reports", True)
    If blnCreate Then
        Call gobjFile.CreateFolder(App.Path & "\Standard")
        Call gobjFile.CreateFolder(App.Path & "\NewStandard")
        Call gobjFile.CreateFolder(App.Path & "\ThisProcedure")
        Call gobjFile.CreateFolder(App.Path & "\Reports")
    End If
End Sub

Private Function LoadSQLInfo(ByVal objSQL As clsSQLInfo) As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    If mlngKey <> 0 Then
        If UCase(objSQL.BlockName) <> UCase(cboProcName.Text) Then
            MsgBox "�洢�������Ʋ�ƥ�䣬ѡ��Ĺ���Ϊ""" & objSQL.BlockName & """��", vbInformation, Me.Caption
            Exit Function
        End If
    End If

    strSQL = "Select a.Id From Zlprocedure a Where Upper(a.����)= [1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ����", UCase(objSQL.BlockName))
    If Not rsTmp.EOF Then
        mlngKey = Val(rsTmp!Id & "")
        Call Form_Load
    Else
        cboProcName.Locked = False: cboOwner.Locked = False
        mrsProcedure.Filter = "Object_Name='" & UCase(objSQL.BlockName) & "'"
        If mrsProcedure.EOF And mptType <> ProcType.�û����� Then
            MsgBox "���ݿ��в����ڴ洢���̻���""" & objSQL.BlockName & """", vbInformation, Me.Caption
            Exit Function
        End If
        cboProcName.Text = objSQL.BlockName
        If Not mrsProcedure.EOF Then
            cboOwner.Text = mrsProcedure!Owner & ""
            cboProcName.Locked = True
            cboOwner.Locked = True
        End If
    End If
    synProcEdit.Text = objSQL.SQL
    LoadSQLInfo = True
End Function

Private Function GetCurrentProctext(Optional ByVal blnAppEnd As Boolean) As String
'���ܣ���ȡ��ǰ�༭���Ĺ�������
    Dim strSQL As String, blnHaveEnd As Boolean
    Dim i As Long

    For i = 1 To synProcEdit.RowsCount
        strSQL = strSQL & IIf(strSQL = "", "", vbCrLf) & synProcEdit.RowText(i)
        If blnAppEnd Then
            If TrimComment(TrimEx(synProcEdit.RowText(i))) = "/" Then
                blnHaveEnd = True
            End If
        End If
    Next
    'û���������������Զ�����һ����
    If Not blnHaveEnd And blnAppEnd Then
        strSQL = strSQL & IIf(strSQL = "", "", vbCrLf) & "/"
    End If
    GetCurrentProctext = strSQL
End Function

Private Function CompareProcPars(ByVal rsLeft As ADODB.Recordset, ByVal rsRigth As ADODB.Recordset) As String
'���ܣ��Դ洢���̽��бȽϣ����رȽϽ�������޲��췵�ؿ�
'������strLeftInfo=��ߴ洢���̵Ĳ�����Ϣ
'      strRightInfo=�ұ߹��̵Ĵ洢������Ϣ
'���أ�������Ϣ���޲��첻���ء�
    Dim rsCom As ADODB.Recordset, strErr As String, intIndex As Integer
    Dim strSQL As String, rsDataType As ADODB.Recordset, strTMp As String
    Dim arrTmp As Variant
    
    
    On Error GoTo errH
    If gobjFile.FileExists("C:\rsLeft.xml") Then Call gobjFile.DeleteFile("C:\rsLeft.xml", True)
    If gobjFile.FileExists("C:\rsRigth.xml") Then Call gobjFile.DeleteFile("C:\rsRigth.xml", True)
    If gobjFile.FileExists("C:\rsCom.xml") Then Call gobjFile.DeleteFile("C:\rsCom.xml", True)
    rsLeft.Save "C:\rsLeft.xml", adPersistXML
    rsRigth.Save "C:\rsRigth.xml", adPersistXML
    '-1ɾ����0-����,1-����,2-����
    Set rsCom = GetCompareRec(rsLeft, rsRigth, "λ��", "����,��������,����,Ĭ��ֵ", "����,λ��")
    rsCom.Save "C:\rsCom.xml", adPersistXML
    With rsCom
        '�鿴���ƣ������Ƿ�ı�
        .Filter = "MainKey='-1'"
        If !State = 2 Or !���� <> !����_New Then
            intIndex = intIndex + 1
            strErr = strErr & vbNewLine & intIndex & "-�������ͻ����Ʋ���:��" & !���� & "��" & !���� & " <---> ��" & !����_New & "��" & !����_New
        End If
        .Filter = "State<>0"
        If .RecordCount = 0 Then Exit Function '�޲��죬���˳�
        '�ȽϷ���ֵ
        .Filter = "MainKey='0' And State <> 0"
        If .RecordCount > 0 Then
            intIndex = intIndex + 1
            strErr = strErr & vbNewLine & intIndex & "-����ֵ���Ͳ���:" & IIf(!State = 1, "�޷�������", !����) & " <---> " & IIf(!State = -1, "�޷�������", !����_New)
        End If
        '�Ƚϲ���,���ȴ������ȱʧ��������
        .Filter = "MainKey<>'0' And MainKey<>'-1' And State=-1" 'ȱʧ�����Ƚ�
        .Sort = "MainKey"
        Do While Not .EOF
            intIndex = intIndex + 1
            strErr = strErr & vbNewLine & intIndex & "-��" & !MainKey & "λ����ȱʧ:����:" & !���� & " �������:" & !���� & " ��������:" & !�������� & !���� & IIf(!Ĭ��ֵ & "" = "", "", " Ĭ��ֵ:" & !Ĭ��ֵ)
            .MoveNext
        Loop
        .Filter = "MainKey<>'0' And MainKey<>'-1' And State=1" '��Ӳ���
        .Sort = "MainKey"
        Do While Not .EOF
            intIndex = intIndex + 1
            strErr = strErr & vbNewLine & intIndex & "-������" & !MainKey & "λ����:����:" & !����_New & " �������:" & !����_New & " ��������:" & !��������_New & !����_New & IIf(!Ĭ��ֵ_New & "" = "", "", " Ĭ��ֵ:" & !Ĭ��ֵ_New)
            .MoveNext
        Loop
        '�������ͱ��
        '�ȴ���̬���͵Ĳ���
        .Filter = "(MainKey<>'0' And MainKey<>'-1' And State=2 And ��������<>'') OR (MainKey<>'0' And MainKey<>'-1' And State=2 And ��������_New<>'')"
        .Sort = "MainKey"
        Do While Not .EOF
            If !�������� & "" <> "" Then
                If InStr(strSQL & ";", ";" & !�������� & ";") = 0 Then
                    strSQL = strSQL & ";" & !��������
                End If
            End If
            If !��������_New & "" <> "" Then
                If InStr(strSQL & ";", ";" & !��������_New & ";") = 0 Then
                    strSQL = strSQL & ";" & !��������_New
                End If
            End If
            .MoveNext
        Loop
        If strSQL <> "" Then
            strSQL = Mid(strSQL, 2)
            strSQL = "Select a.C1 || '.' || a.C2 Key, b.Data_Type" & vbNewLine & _
                    "From (Select C1, C2 From Table(f_Str2list2('" & strSQL & "', ';', '.'))) a, All_Tab_Columns b" & vbNewLine & _
                    "Where a.C1 = b.Table_Name(+) And a.C2 = b.Column_Name(+)"
            Set rsDataType = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
            .Sort = "MainKey"
            Do While Not .EOF
                If !�������� & "" <> "" Then
                    rsDataType.Filter = "Key='" & !�������� & "'"
                    .Update "����", rsDataType!DATA_TYPE & ""
                End If
                If !��������_New & "" <> "" Then
                    rsDataType.Filter = "Key='" & !��������_New & "'"
                    .Update "����_New", rsDataType!DATA_TYPE & ""
                End If
                .MoveNext
            Loop
        End If
        '�����ȽϽ��
        .Filter = "MainKey<>'0' And MainKey<>'-1' And State=2"
        .Sort = "MainKey"
        Do While Not .EOF
            '�����������ڲ���������Ͳ��죬��Զ��߽���
            If !���� <> "" And !����_New <> "" Then
                If !���� & "" = !����_New & "" Then
                    strTMp = !DifInfo & ""
                    If InStr("," & strTMp & ",", ",��������,") > 0 Then
                        strTMp = Replace("," & strTMp, ",��������", "")
                    End If
                    If InStr("," & strTMp & ",", ",����,") > 0 Then
                        strTMp = Replace("," & strTMp, ",����", "")
                    End If
                    .Update "DifInfo", strTMp
                End If
            End If
            If !DifInfo & "" = "" Then
                .Update "State", 0
            End If
            .MoveNext
        Loop
        .Filter = "MainKey<>'0' And MainKey<>'-1' And State=2"
        .Sort = "MainKey"
        Do While Not .EOF
            intIndex = intIndex + 1
            strErr = strErr & vbNewLine & intIndex & "-��" & !MainKey & "λ�������ڲ���(�������Ʋ���):" & vbNewLine & _
                     "����:" & !���� & " �������:" & !���� & " ��������:" & IIf(!�������� = "", !����, !�������� & "%TYPE(" & IIf(!���� = "", "�޷���ȡ����", !����) & ")") & IIf(!Ĭ��ֵ & "" = "", "", " Ĭ��ֵ:" & !Ĭ��ֵ) & vbNewLine & _
                     "����:" & !����_New & " �������:" & !����_New & " ��������:" & IIf(!��������_New = "", !����_New, !��������_New & "%TYPE(" & IIf(!����_New = "", "�޷���ȡ����", !����_New) & ")") & IIf(!Ĭ��ֵ_New & "" = "", "", " Ĭ��ֵ:" & !Ĭ��ֵ_New)
            .MoveNext
        Loop
    End With
    CompareProcPars = strErr
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function
