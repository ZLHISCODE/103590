VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPurchasePlan 
   Caption         =   "�ɹ��ƻ�����"
   ClientHeight    =   7680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10890
   Icon            =   "frmPurchasePlan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   10890
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picGetParams 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   360
      ScaleHeight     =   3855
      ScaleWidth      =   3855
      TabIndex        =   3
      Top             =   2880
      Width           =   3855
      Begin VB.Frame fraParams 
         Height          =   3015
         Left            =   180
         TabIndex        =   17
         Top             =   120
         Width           =   3015
         Begin VB.CheckBox chkUpload 
            Caption         =   "�����Ѿ�����ļ�¼"
            Height          =   180
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtParam02 
            Height          =   270
            Left            =   1440
            TabIndex        =   11
            Top             =   1680
            Width           =   1335
         End
         Begin VB.OptionButton optParams01 
            Caption         =   "�������(&C)"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   2040
            Width           =   1290
         End
         Begin VB.OptionButton optParams01 
            Caption         =   "�ɹ�����(&R)"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   1320
            Value           =   -1  'True
            Width           =   1290
         End
         Begin VB.TextBox txtParam01 
            Height          =   270
            Left            =   1440
            TabIndex        =   9
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton cmdPS 
            Caption         =   "��"
            Height          =   255
            Left            =   2520
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox txtProvider 
            Height          =   270
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker dtpParam01 
            Height          =   270
            Left            =   1440
            TabIndex        =   13
            Top             =   2040
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            Format          =   284491777
            CurrentDate     =   40290
         End
         Begin MSComCtl2.DTPicker dtpParam02 
            Height          =   270
            Left            =   1440
            TabIndex        =   15
            Top             =   2400
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            Format          =   280428545
            CurrentDate     =   40290
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   2
            Left            =   1200
            TabIndex        =   10
            Top             =   1680
            Width           =   180
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   3
            Left            =   1200
            TabIndex        =   14
            Top             =   2400
            Width           =   180
         End
         Begin VB.Label lblProvider 
            AutoSize        =   -1  'True
            Caption         =   "��Ӧ��(&P)"
            Height          =   180
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   810
         End
      End
      Begin VB.CommandButton cmdGetData 
         Caption         =   "��ȡ����(&G)"
         Height          =   375
         Left            =   180
         TabIndex        =   16
         Top             =   3240
         Width           =   1215
      End
   End
   Begin VB.PictureBox picView 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   2760
      ScaleHeight     =   1575
      ScaleWidth      =   3300
      TabIndex        =   0
      Top             =   1080
      Width           =   3300
      Begin VSFlex8Ctl.VSFlexGrid vsfView 
         Height          =   1000
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2655
         _cx             =   4683
         _cy             =   1764
         Appearance      =   1
         BorderStyle     =   1
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483645
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      End
   End
   Begin MSComctlLib.TreeView tvwProvider 
      Height          =   1215
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2143
      _Version        =   393217
      Indentation     =   529
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   7320
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   635
      SimpleText      =   $"frmPurchasePlan.frx":1CFA
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPurchasePlan.frx":1D41
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14129
            Text            =   "��ɫ��Ϊ�Ѵ���������ݣ� ��ɫ��Ϊ����ѡ������ݣ� ��ɫ��Ϊ�������ݡ�"
            TextSave        =   "��ɫ��Ϊ�Ѵ���������ݣ� ��ɫ��Ϊ����ѡ������ݣ� ��ɫ��Ϊ�������ݡ�"
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
   Begin XtremeCommandBars.CommandBars cmbMain 
      Left            =   8520
      Top             =   1200
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPurchasePlan.frx":25D5
      Left            =   8040
      Top             =   1200
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPurchasePlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case enm_Pop_File.FilePrintSet
            frmOutsideLinkSet.Show vbModal, Me
        Case enm_Pop_File.EditProcess
            Call ProcProcess
        Case enm_Pop_File.EditCurrChoose
            SignData vsfView, 4, True
        Case enm_Pop_File.EditCurrCancel
            SignData vsfView, 4, False
        Case enm_Pop_File.EditChooChoose
            SignData vsfView, 3, True
        Case enm_Pop_File.EditChooCancel
            SignData vsfView, 3, False
        Case enm_Pop_File.EditAllChoose
            SignData vsfView, 1, True
        Case enm_Pop_File.EditAllCancel
            SignData vsfView, 0, False
        Case enm_Pop_File.ViewRefresh
            Call cmdGetData_Click
        Case enm_Pop_File.ViewFindButton
            Call FindString
        Case enm_Pop_File.ViewToolsButton
            Control.Checked = Not Control.Checked
            cmbMain(2).Visible = Control.Checked
            cmbMain.RecalcLayout
        Case enm_Pop_File.ViewToolsLabel
            Dim cbcControl As CommandBarControl
            Control.Checked = Not Control.Checked
            For Each cbcControl In Me.cmbMain(2).Controls
                cbcControl.Style = IIf(cbcControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cmbMain.RecalcLayout
        Case enm_Pop_File.ViewToolsIcon
            Control.Checked = Not Control.Checked
            cmbMain.Options.LargeIcons = Not Me.cmbMain.Options.LargeIcons
            cmbMain.RecalcLayout
        Case enm_Pop_File.ViewStatebar
            Control.Checked = Not Control.Checked
            stbThis.Visible = Not stbThis.Visible
            cmbMain.RecalcLayout
        Case enm_Pop_File.FileExit
            Unload Me
    End Select
End Sub

Private Sub cmbMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cmdGetData_Click()
    Dim strDB As String, strServer As String, strUser As String, strPWD As String
    Dim strSQL As String, strProvider As String
    Dim isConn As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim dtEnd As Date

    If optParams01(0).Value Then
        If Len(Trim(txtParam01.Text)) = 0 Or Len(Trim(txtParam02.Text)) = 0 Then
            MsgBox "������Ҫ��ȡ[�ɹ�����]��ʼ����������Ϣ��", vbInformation, GSTR_MESSAGE
            txtParam01.SetFocus
            Exit Sub
        End If
    Else
        If Len(Trim(dtpParam01.Value)) = 0 Or Len(Trim(dtpParam02.Value)) = 0 Then
            MsgBox "������Ҫ��ȡ[�������]��ʼ����������Ϣ��", vbInformation, GSTR_MESSAGE
            dtpParam01.SetFocus
            Exit Sub
        End If
        If IsDate(dtpParam01.Value) = False Or IsDate(dtpParam02.Value) = False Then
            dtpParam01.SetFocus
            MsgBox "���������[�������]��", vbInformation, GSTR_MESSAGE
            Exit Sub
        End If
    End If

'��ȡ�ⲿ����
'step1 �����ⲿ���ݿ�
    strDB = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="DBNAME", Default:="")
    strServer = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="SERVER", Default:="")
    strUser = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="USER", Default:="")
    strPWD = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="PASSWORD", Default:="")
    strPWD = StringEnDeCodecn(strPWD, 68)
    'Ĭ��MSSQL��ʽ����
    isConn = MSSQLServerOpen(strServer, strDB, strUser, strPWD)
    
    If isConn = False Then
        MsgBox "���ӷ�����ʧ�ܣ��������м����ݿ�����ӣ�", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If

'step2 ��ȡ���ݼ�
    Screen.MousePointer = vbHourglass
    
    strProvider = Trim(txtProvider.Text)
    
    On Error GoTo ErrHand
    strSQL = "select a.id,a.no,b.���,a.�������,a.��������,null CSTCODE,null CSTNAME,b.ҩƷid,c.����,null GOODENAME" _
           & "  ,null GOODGNAME,c.���,c.ҩ�ⵥλ,c.ҩ���װ PACKNUM,e.id ��Ӧ��id,b.�ϴ�������,a.�ⷿid ҩ��ID,d.���� ҩ��" _
           & "  ,a.ҩ��ID, f.���� ҩ��, b.�ƻ�����/c.ҩ���װ �ƻ�����" _
           & "  ,b.����*c.ҩ���װ ����,null MARK,b.�ϴι�Ӧ��,null ImportDate,null ImportUserId,null ImportUserFlag,null ToCode, b.�Ƿ��ϴ� " _
           & "from ҩƷ�ɹ��ƻ� a, ҩƷ�ƻ����� b, ҩƷĿ¼ c, ���ű� d, ��Ӧ�� e, ���ű� f " _
           & "where a.id=b.�ƻ�id and b.ҩƷid=c.ҩƷid and a.�ⷿid=d.id(+) and a.ҩ��id=f.id(+) and b.�ϴι�Ӧ��=e.����(+) " _
           & "  and Nvl(B.�ƻ�����, 0) > 0 And Nvl(C.ҩ���װ, 0) > 0" _
           & "  and (d.����ʱ�� is null or d.����ʱ��=to_date('3000-1-1', 'yyyy-mm-dd')) " _
           & "  and (e.����ʱ�� is null or e.����ʱ��=to_date('3000-1-1', 'yyyy-mm-dd')) "
    '��Ӧ������
    If strProvider <> "" Then 'And strProvider <> "[ȫ��]" Then
        strSQL = strSQL & " and b.�ϴι�Ӧ�� like '%" & strProvider & "%'"
    End If
    '�����Ѿ��ϴ�
    If chkUpload.Value = False Then
        strSQL = strSQL & " and nvl(b.�Ƿ��ϴ�,0)=0 "
    End If
    If optParams01(0).Value Then        '�ɹ�����
        strSQL = strSQL & " and a.no between [1] and [2] order by " & IIf(chkUpload.Value, "b.�Ƿ��ϴ�,", "") & "a.no,b.���"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtParam01.Text, txtParam02.Text)
    Else                                '�������
        strSQL = strSQL & " and a.������� between [1] and [2] order by " & IIf(chkUpload.Value, "b.�Ƿ��ϴ�,", "") & "a.no,b.���"
        dtEnd = dtpParam02.Value & " 23:59:59"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtpParam01.Value, dtEnd)
    End If
    
    If rsTmp.RecordCount <= 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "ZLHIS���ݿ�����ʱ�����ݿɻ�ȡ��", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    
'step3 װ������
    DataLoading vsfView, rsTmp, 0
    RefreshTVWProvider tvwProvider, vsfView
    Screen.MousePointer = vbDefault
    'MsgBox "��ȡ������ɣ�", vbInformation, GSTR_MESSAGE
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    MsgBox "�����쳣����", vbCritical, GSTR_MESSAGE
End Sub

Private Sub cmdPS_Click()
    ProviderSelecter Me, txtProvider, True
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1: Item.Handle = picView.hwnd
        Case 2: Item.Handle = tvwProvider.hwnd
        Case 3: Item.Handle = picGetParams.hwnd
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub Form_Load()
    InitCommandBars cmbMain
    Call InitDKPMain
    Call InitToolBar
    Call SetMenu
    InitVSF vsfView, False
    dtpParam01.Value = Date - 7
    dtpParam02.Value = Date
    optParams01_Click 0
End Sub

Private Sub InitDKPMain()
'��ʼ��dkpMain
    Dim pneMain As Pane, pneProvider As Pane, pneGetParams As Pane, pneFind As Pane
    With dkpMain
        Set pneMain = .CreatePane(1, Me.ScaleHeight, 0, DockRightOf)
        pneMain.Options = PaneNoCloseable + PaneNoHideable + PaneNoFloatable
        pneMain.Title = "����������"
        
        Set pneProvider = .CreatePane(2, 230, 400, DockLeftOf)
        pneProvider.Options = PaneNoCloseable + PaneNoFloatable '+ PaneNoHideable
        pneProvider.Title = "��Ӧ���б�"
        pneProvider.MinTrackSize.Width = 230
        pneProvider.MinTrackSize.Height = 50
        
        Set pneGetParams = .CreatePane(3, 230, 250, DockBottomOf, pneProvider)
        pneGetParams.Options = PaneNoCloseable + PaneNoFloatable
        pneGetParams.Title = "��������"
        pneGetParams.MinTrackSize.Height = 250
        pneGetParams.MaxTrackSize.Height = 250
        pneGetParams.MinTrackSize.Width = 230
        
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        If Not cmbMain Is Nothing Then .SetCommandBars cmbMain
    End With
    
End Sub

Private Sub InitToolBar()
    Dim cbcControl As CommandBarControl
    Dim cbrToolBar As CommandBar

    Set cbrToolBar = cmbMain.Add("������", xtpBarTop)
    'cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.EnableDocking xtpFlagAlignTop
    With cbrToolBar.Controls
        'Set cbcControl = .Add(xtpControlButton, arrMenuBars(1).Id, arrMenuBars(1).Caption)
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.FilePrintSet, "����")
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditProcess, "����")
        cbcControl.BeginGroup = True
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.ViewRefresh, "ˢ��")
        cbcControl.BeginGroup = True
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.FileExit, "�˳�")
        cbcControl.BeginGroup = True
    End With
    For Each cbcControl In cbrToolBar.Controls
        If cbcControl.Type = xtpControlButton Then
            cbcControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
'    Set cbrToolBar = cmbMain.Add("����", xtpBarTop)
'    cbrToolBar.EnableDocking xtpFlagAlignTop
'    With cbrToolBar.Controls
'        Set cbcControl = .Add(xtpControlLabel, enm_Pop_File.ViewFindTitle, "����(��Ʊ��)��")
'        Set cbcControl = .Add(xtpControlEdit, enm_Pop_File.ViewFindEdit, "")
'        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.ViewFindButton, "")
'    End With
    
End Sub


Private Sub SetMenu()
    Dim cbcControl As CommandBarControl, cbcControlParent As CommandBarControl
    Dim cbpMenuBar As CommandBarPopup
    
    cmbMain.ActiveMenuBar.Title = "�˵�"
    cmbMain.ActiveMenuBar.EnableDocking xtpFlagAlignTop
    
    Set cbpMenuBar = cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enm_Pop_File.File, "�ļ�(&F)", -1, False)
    cbpMenuBar.Id = enm_Pop_File.File
    With cbpMenuBar.CommandBar.Controls
        'Set cbcControl = .Add(xtpControlButton, arrMenuBars(1).Id, arrMenuBars(1).Caption & arrMenuBars(1).HotKey)
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.FilePrintSet, "�������ݿ�����(&S)")
        
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.FileExit, "�˳�(&X)")
        cbcControl.BeginGroup = True        '����Ϊһ��Ŀ�ʼ
    End With
    
    Set cbpMenuBar = cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enm_Pop_File.Edit, "�༭(&E)", -1, False)
    cbpMenuBar.Id = enm_Pop_File.Edit
    With cbpMenuBar.CommandBar.Controls
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditProcess, "���ݴ���(&P)")
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditCurrChoose, "��ǰ��Ӧ�̴�")
        
        cbcControl.BeginGroup = True
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditCurrCancel, "��ǰ��Ӧ��ȡ��")
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditChooChoose, "ѡ�д�")
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditChooCancel, "ѡ��ȡ��")
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditAllChoose, "ȫ����")
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditAllCancel, "ȫ��ȡ��")
    End With
    
    Set cbpMenuBar = cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enm_Pop_File.View, "�鿴(&V)", -1, False)
    cbpMenuBar.Id = enm_Pop_File.View
    With cbpMenuBar.CommandBar.Controls
        Set cbcControlParent = .Add(xtpControlPopup, enm_Pop_File.ViewTools, "������(&T)")
        Set cbcControl = cbcControlParent.CommandBar.Controls.Add(xtpControlButton, enm_Pop_File.ViewToolsButton, "��׼��ť(&S)", -1, False)
        cbcControl.Checked = True
        Set cbcControl = cbcControlParent.CommandBar.Controls.Add(xtpControlButton, enm_Pop_File.ViewToolsLabel, "�ı���ǩ(&T)", -1, False)
        cbcControl.Checked = True
        Set cbcControl = cbcControlParent.CommandBar.Controls.Add(xtpControlButton, enm_Pop_File.ViewToolsIcon, "��ͼ��(&B)", -1, False)
        cbcControl.Checked = True
        
        Set cbcControlParent = .Add(xtpControlButton, enm_Pop_File.ViewStatebar, "״̬��(&S)")
        cbcControlParent.Checked = True
        
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.ViewRefresh, "ˢ��(&R)")
        cbcControl.ShortcutText = "F5"
        cbcControl.BeginGroup = True
    End With
    
    Set cbpMenuBar = cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enm_Pop_File.Help, "����(&H)", -1, False)
    cbpMenuBar.Id = enm_Pop_File.Help
    With cbpMenuBar.CommandBar.Controls
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.HelpHelp, "��������(&H)")
        Set cbcControl = .Add(xtpControlPopup, enm_Pop_File.HelpWeb, "&WEB�ϵ�����")
        cbcControl.CommandBar.Controls.Add xtpControlButton, enm_Pop_File.HelpWebhome, "������ҳ(&H)", -1, False
        cbcControl.CommandBar.Controls.Add xtpControlButton, enm_Pop_File.HelpWebBBS, "������̳(&F)", -1, False
        cbcControl.CommandBar.Controls.Add xtpControlButton, enm_Pop_File.HelpWebFeelback, "���ͷ���(&M)", -1, False
        
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.HelpAbout, "����(&A)��")
        cbcControl.BeginGroup = True
    End With
    
    '�����
    With cmbMain.KeyBindings
'        .Add FCONTROL, Asc("X"), conMenu_File_Exit
        .Add 0, VK_F5, enm_Pop_File.ViewRefresh
        .Add 0, VK_F1, enm_Pop_File.HelpHelp
    End With
    
    For Each cbcControl In cbpMenuBar.Controls
        cbcControl.Style = xtpButtonIconAndCaption
    Next

End Sub

Private Sub optParams01_Click(Index As Integer)
    Dim lngBackColor As Long
    On Error Resume Next
    If Index = 0 Then
        txtParam01.Enabled = True
        txtParam02.Enabled = True
        txtParam01.BackColor = vbWhite
        txtParam02.BackColor = vbWhite
        dtpParam01.Enabled = False
        dtpParam02.Enabled = False
        txtParam01.SetFocus
    Else
        txtParam01.Enabled = False
        txtParam02.Enabled = False
        txtParam01.BackColor = &H80000004
        txtParam02.BackColor = &H80000004
        dtpParam01.Enabled = True
        dtpParam02.Enabled = True
        dtpParam01.SetFocus
    End If
End Sub

Private Sub picGetParams_Resize()
    fraParams.Width = IIf(picGetParams.Width > 300, picGetParams.Width - 300, 0)
    txtProvider.Width = IIf(picGetParams.Width > 700 + cmdPS.Width, picGetParams.Width - 700 - cmdPS.Width, 0)
    cmdPS.Left = IIf(txtProvider.Width > 0, txtProvider.Left + txtProvider.Width + 20, 0)
    'fraParams01.Width = IIf(picGetParams.Width > fraParams01.Left + 500, picGetParams.Width - fraParams01.Left - 500, 0)
End Sub

Private Sub picView_Resize()
    With vsfView
        .Top = 0
        .Left = 0
        .Width = picView.Width
        .Height = picView.Height
    End With
End Sub

Private Sub tvwProvider_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer, intCounter As Integer
    Dim bytState As Byte
    'Check״̬��ʾ
    vsfView.Redraw = flexRDNone
    If Node.Key = "Root" Then
        For i = 2 To tvwProvider.Nodes.Count
            tvwProvider.Nodes(i).Checked = Node.Checked
        Next
    Else
        For i = 2 To tvwProvider.Nodes.Count
            If i = 2 Then
                If tvwProvider.Nodes(i).Checked Then
                    bytState = 2
                Else
                    bytState = 1
                End If
            Else
                If (bytState = 1 And tvwProvider.Nodes(i).Checked) Or (bytState = 2 And tvwProvider.Nodes(i).Checked = False) Then
                    bytState = 0
                    Exit For
                End If
            End If
        Next
        Select Case bytState
            Case 1: tvwProvider.Nodes(1).Checked = False
            Case 2: tvwProvider.Nodes(1).Checked = True
            Case Else: tvwProvider.Nodes(1).Checked = 0
        End Select
    End If
    '����VSFView����ɵļ�¼
    If Node.Key = "Root" Then
        For i = 1 To vsfView.Rows - 1
            vsfView.RowHidden(i) = Not Node.Checked
        Next
    Else
        For i = 1 To vsfView.Rows - 1
            If Node.Tag = -1 Then
                If vsfView.TextMatrix(i, vsfView.ColIndex("imported")) = "0,0" Then
                    vsfView.RowHidden(i) = Not Node.Checked
                End If
            ElseIf Node.Tag = Val(vsfView.TextMatrix(i, vsfView.ColIndex("providerid"))) Then
                If vsfView.TextMatrix(i, vsfView.ColIndex("imported")) <> "0,0" Then
                    vsfView.RowHidden(i) = Not Node.Checked
                End If
            End If
            
        Next
    End If
    '��д���
    intCounter = 1
    For i = 1 To vsfView.Rows - 1
        If vsfView.RowHidden(i) = False Then
            vsfView.TextMatrix(i, 1) = intCounter
            intCounter = intCounter + 1
        End If
    Next
    vsfView.Redraw = flexRDBuffered
End Sub

Private Sub txtProvider_GotFocus()
    txtProvider.SelStart = 0: txtProvider.SelLength = Len(txtProvider.Text)
End Sub

Private Sub txtProvider_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ProviderSelecter(Me, txtProvider, False)
    End If
End Sub

Private Sub vsfView_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfView
        'ѡ������޸�
        If Col = .ColIndex("choose") Then
            If Mid(.TextMatrix(Row, .ColIndex("imported")), 3, 1) = "1" Then
                Cancel = False
            Else
                Cancel = True
            End If
        ElseIf Col = .ColIndex("qty") Then Cancel = False
        ElseIf Col = .ColIndex("remark") Then Cancel = False
        Else: Cancel = True
        End If
    End With
End Sub

Private Sub vsfView_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < 3 Then Cancel = True
End Sub

Private Sub vsfView_EnterCell()
    With vsfView
        '������ɫ
        .ForeColorSel = .Cell(flexcpForeColor, .Row, 3)
    End With
End Sub

Private Sub vsfView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopupMenu As CommandBarPopup
    Dim cbcControl As CommandBarControl
    
    If vsfView.Rows <= 1 Then Exit Sub
    
    If Button = vbRightButton Then
        Set objPopupMenu = cmbMain.ActiveMenuBar.FindControl(, enm_Pop_File.Edit)
        If Not objPopupMenu Is Nothing Then
            '����Ҫ���صĲ˵���
            For Each cbcControl In objPopupMenu.CommandBar.Controls
                If cbcControl.Id = enm_Pop_File.EditProcess Then
                    cbcControl.Visible = False
                    Exit For
                End If
            Next
            objPopupMenu.CommandBar.ShowPopup
            '�ָ�
            If Not cbcControl Is Nothing Then
                cbcControl.Visible = True
            End If
        End If
    End If
End Sub

Private Sub vsfView_RowColChange()
    '��ǰ��¼�ü�ͷָʾ
    vsfView.Cell(flexcpText, 0, 0, vsfView.Rows - 1, 0) = ""
    If vsfView.Row > 0 Then
        vsfView.Cell(flexcpFontName, , 0) = "Marlett"
        vsfView.TextMatrix(vsfView.Row, 0) = 4
    End If
End Sub

Private Sub FindString()
    Dim cbeFind As CommandBarEdit
    Set cbeFind = cmbMain.FindControl(, enm_Pop_File.ViewFindEdit)
    
    If cbeFind Is Nothing Then Exit Sub
    
    If Trim(cbeFind.Text) <> "" And vsfView.Rows > 1 Then
        '���ҷ�Ʊ��
        Dim i As Integer
        With vsfView
            For i = 1 To .Rows - 1
                If UCase(.TextMatrix(i, .ColIndex("invoice"))) = UCase(Trim(cbeFind.Text)) And .RowHidden(i) = False Then
                    .Row = i
                    .TopRow = i
                    .SetFocus
                    Exit Sub
                End If
            Next
        End With
        MsgBox "δ�ҵ���¼��ķ�Ʊ�ţ�", , GSTR_MESSAGE
    End If
End Sub

Private Sub ProcProcess()
    Dim strTmp As String
    
    If vsfView.Rows <= 1 Or CheckRecord(vsfView) = False Then
        MsgBox "�����ݿ��Դ������Ȼ�ȡ���ݣ�", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If

    '�ⲿ���ݿ��Ƿ�����
    On Error GoTo ExitSub
    If gcnOutside.State = adStateClosed Then gcnOutside.Open
    On Error GoTo 0

    '�������ݿ�
    If MsgBox("��ȷ��Ҫ������", vbInformation Or vbYesNo Or vbDefaultButton2, GSTR_MESSAGE) = vbNo Then Exit Sub
    
    Call ProcExport
    
    Exit Sub
    
ExitSub:
    MsgBox "�ⲿ���ݿ�����ʧ��!", vbCritical
    Exit Sub
End Sub

Private Sub ProcExport()
    '�ƻ����������ݴ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMess As String
    Dim i As Long, intReturn As Long
    
    '����޹�Ӧ�̵�����
    If CheckRowProvider(i) = False Then
        MsgBox "��" & i & "���޹�Ӧ����Ϣ��", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    
    With vsfView
        gcnOutside.BeginTrans
        gcnOracle.BeginTrans
        On Error GoTo ErrHand
        For i = 1 To .Rows - 1
            '��������
            If Val(vsfView.ValueMatrix(i, vsfView.ColIndex("choose"))) = 0 Or vsfView.RowHidden(i) = True Then GoTo ProcEnd
            
            strSQL = "declare @i_rtn int, @s_msg varchar(200) " & Chr(13)
            strSQL = strSQL & "execute sj_insertBill_pro " _
                   & " '" & .TextMatrix(i, .ColIndex("planno")) & "'" _
                   & ",'" & .TextMatrix(i, .ColIndex("xh")) & "'" _
                   & ",'" & .TextMatrix(i, .ColIndex("cdate")) & "'" _
                   & ",'" & .TextMatrix(i, .ColIndex("edate")) & "'" _
                   & ", null, null" _
                   & ",'" & .TextMatrix(i, .ColIndex("id")) & "'" _
                   & ",'" & .TextMatrix(i, .ColIndex("name")) & "'" _
                   & ",null, null" _
                   & ",'" & .TextMatrix(i, .ColIndex("spec")) & "'" _
                   & ",'" & .TextMatrix(i, .ColIndex("unit")) & "'" _
                   & ",null" _
                   & ",'" & .TextMatrix(i, .ColIndex("producer")) & "'" _
                   & ",'" & .TextMatrix(i, .ColIndex("wh_id")) _
                   & "|" & IIf(Val(.TextMatrix(i, .ColIndex("dh_id"))) = 0, .TextMatrix(i, .ColIndex("wh_id")), .TextMatrix(i, .ColIndex("dh_id"))) & "'" _
                   & ",'" & IIf(Val(.TextMatrix(i, .ColIndex("dh_id"))) = 0, .TextMatrix(i, .ColIndex("wh")), .TextMatrix(i, .ColIndex("dh"))) & "'" _
                   & "," & .TextMatrix(i, .ColIndex("qty")) _
                   & "," & Round(.TextMatrix(i, .ColIndex("price")), 2) _
                   & ",'" & .TextMatrix(i, .ColIndex("remark")) & "'" _
                   & ",'" & IIf(.TextMatrix(i, .ColIndex("providerid")) = "" _
                     , .TextMatrix(i, .ColIndex("provider")) _
                     , .TextMatrix(i, .ColIndex("providerid"))) & "'" _
                   & ",null,null,null,null,null,@i_rtn output, @s_msg output " & Chr(13)
            strSQL = strSQL & "select @i_rtn i_rtn, @s_msg s_msg "
            rsTmp.Open strSQL, gcnOutside
            If rsTmp.EOF Then
                intReturn = 0
                strMess = ""
            Else
                intReturn = rsTmp!i_rtn
                strMess = rsTmp!s_msg
            End If
            rsTmp.Close
            
            '�ƻ����ݱ���ϴ�
            strSQL = "zl_ҩƷ�ƻ�����_Upload(" & .TextMatrix(i, .ColIndex("planid")) _
                   & "," & .TextMatrix(i, .ColIndex("xh")) & ")"
            gobjComLib.zlDatabase.ExecuteProcedure strSQL, Me.Caption & "-��Ǽƻ��ϴ�"
            
            .TextMatrix(i, .ColIndex("mess")) = strMess
            If intReturn = 1 Then
                .TextMatrix(i, .ColIndex("mess")) = "OK"
            End If
            
ProcEnd:
        Next

        gcnOracle.CommitTrans
        gcnOutside.CommitTrans
        '�����Ѿ����������
        For i = .Rows - 1 To 1 Step -1
            If .TextMatrix(i, .ColIndex("mess")) = "OK" Then
                .RemoveItem i
            Else
                If .Cell(flexcpChecked, i, .ColIndex("choose")) = Checked And InStr(.TextMatrix(i, .ColIndex("mess")), "�Ѵ���") > 0 Then
                    If MsgBox("��" & i & "�������Ѿ��е��������Ƿ�����鿴��ʾ��", vbInformation + vbYesNo, GSTR_MESSAGE) = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
        Next
    
    End With
    
    Exit Sub

ErrHand:
    gcnOracle.RollbackTrans
    gcnOutside.RollbackTrans
    Call gobjComLib.ErrCenter
End Sub

Private Sub SignData(ByVal vsfVal As VSFlexGrid, ByVal bytVal As Byte, ByVal blnVal As Boolean)
'0: ȫ��ȡ��; 1:ȫ��ѡ��; 2: ѡ��ȡ��; 3:ѡ�д�; 4:��Ӧ��
    Dim i As Integer
    Dim strTmp As String
    
    If vsfVal.Rows < 2 Then Exit Sub
    
    With vsfVal
        strTmp = .TextMatrix(.Row, .ColIndex("provider"))
        'ע��: SelectedRowsҪ��Ч��SelectMode��ҪΪ flexSelectionListBox
        For i = 1 To .Rows - 1
            Select Case bytVal
                Case 0, 1
                    'vsfView.TextMatrix(i, 2) = IIf(blnVal And Mid(vsfView.TextMatrix(i, vsfView.ColIndex("imported")), 3, 1) = "1", "1", "0")
                    .TextMatrix(i, 2) = IIf(blnVal And Right(.TextMatrix(i, .ColIndex("imported")), 1) = "1", "1", "0")
                Case 2, 3
                    If .IsSelected(i) = True Then
                        .TextMatrix(i, 2) = IIf(blnVal And Right(.TextMatrix(i, .ColIndex("imported")), 1) = "1", "1", "0")
                    End If
                Case 4
                    If .TextMatrix(i, .ColIndex("provider")) = strTmp Then
                        .TextMatrix(i, 2) = IIf(blnVal And Right(.TextMatrix(i, .ColIndex("imported")), 1) = "1", "1", "0")
                    End If
            End Select
        Next
    End With
End Sub

Private Function CheckRowProvider(ByRef lngRow As Long) As Boolean
'-----------------------------------------------
'���ܣ���鹩Ӧ������
'������lngRowû�й�Ӧ�����Ƶ��к�
'����ֵ��False���δͨ����True���ͨ��
'-----------------------------------------------
    Dim i As Long

    With vsfView
        For i = 1 To .Rows - 1
            If Val(.ValueMatrix(i, .ColIndex("choose"))) <> 0 And vsfView.RowHidden(i) = False Then
                If Trim(.TextMatrix(i, .ColIndex("provider"))) = "" Then
                    lngRow = i
                    Exit Function
                End If
            End If
        Next
    End With
    CheckRowProvider = True
End Function
