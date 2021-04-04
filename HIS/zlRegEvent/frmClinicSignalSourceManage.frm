VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmClinicSignalSourceManage 
   BorderStyle     =   0  'None
   Caption         =   "�ٴ���Դ����"
   ClientHeight    =   9975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   ScaleHeight     =   9975
   ScaleWidth      =   12045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   9360
      MaxLength       =   100
      TabIndex        =   5
      Top             =   390
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.PictureBox picDetailList 
      BorderStyle     =   0  'None
      Height          =   3075
      Left            =   3930
      ScaleHeight     =   3075
      ScaleWidth      =   3675
      TabIndex        =   3
      Top             =   3390
      Width           =   3675
      Begin zl9RegEvent.ClinicPlanDetailPages CPDPages 
         Height          =   2535
         Left            =   660
         TabIndex        =   4
         Top             =   180
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   4471
         BackColor       =   -2147483628
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
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2445
      Left            =   2790
      ScaleHeight     =   2445
      ScaleWidth      =   6900
      TabIndex        =   0
      Top             =   870
      Width           =   6900
      Begin XtremeReportControl.ReportControl rptData 
         Height          =   1425
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3705
         _Version        =   589884
         _ExtentX        =   6535
         _ExtentY        =   2514
         _StockProps     =   0
         ShowGroupBox    =   -1  'True
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H8000000C&
         Height          =   735
         Left            =   5040
         Top             =   720
         Width           =   405
      End
      Begin XtremeSuiteControls.ShortcutCaption sccTitle 
         Height          =   360
         Left            =   -60
         TabIndex        =   1
         Top             =   -30
         Width           =   7905
         _Version        =   589884
         _ExtentX        =   13944
         _ExtentY        =   635
         _StockProps     =   6
         Caption         =   "��������>�ٴ���Դ����"
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
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin ComctlLib.ImageList imgList16 
      Left            =   7380
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicSignalSourceManage.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicSignalSourceManage.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicSignalSourceManage.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicSignalSourceManage.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicSignalSourceManage.frx":1668
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicSignalSourceManage.frx":1C02
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmClinicSignalSourceManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar�ؼ�
Private mlngModule As Long
Private mstrPrivs As String

Private Enum mRptHeadCol
    COL_ID = 0
    COL_�Ű෽ʽͼ��
    COL_����
    COL_����
    COL_����
    COL_�շ���Ŀ
    COL_ҽ��
    COL_����
    COL_�Ű෽ʽ
    COL_ԤԼ����
    COL_����Ƶ��
    COL_���ջ���
    COL_���տ���״̬
    COL_�ٴ��Ű�
    COL_�����Ա�
    COL_���������
    COL_�Ƿ�ͣ��
    COL_�Ƿ�ɾ��
    COL_����ʱ��
    COL_����ʱ��
End Enum

Private mblnShowStopSignal As Boolean '�Ƿ���ʾ��ͣ�ú�Դ
Private Const conPane_SignalSorceList = 1
Private Const conPane_DetialList = 2
Private mrsWorkTime As ADODB.Recordset
Private mobj���к�����λ  As ������λ���Ƽ�
Private mblnShowDetial As Boolean
Private mlngPreSel��ԴID As Long
Private mintFindType As Integer
Private mrs��Դ As Recordset

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, _
    ByVal strPrivs As String, ByVal lngModule As Long)
    '��ʼ������
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    
    mstrPrivs = strPrivs
    mlngModule = lngModule
End Sub

Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:���˺�
    '����:2016-03-22 14:37:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, strReg As String, panThis As Pane
    Dim panLeft As Pane
    
    Set panLeft = dkpMan.CreatePane(conPane_SignalSorceList, 200, 980, DockLeftOf, Nothing)
    panLeft.Title = "": panLeft.Tag = conPane_SignalSorceList
    panLeft.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panLeft.Handle = picBack.Hwnd
    Set panThis = dkpMan.CreatePane(conPane_DetialList, 100, 280, DockBottomOf, panLeft)
    panThis.Tag = conPane_DetialList
    panThis.Handle = picDetailList.Hwnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    'zlRestoreDockPanceToReg Me, dkpMan, "����"
End Sub

Private Sub LoadDetialData(ByVal lng��ԴId As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ϸ����
    '���:lng��ԴID-��ԴID
    '����:���˺�
    '����:2016-03-22 15:58:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj�����¼�� As �����¼��
    Dim objPan As Pane
    On Error GoTo errHandle
    Set objPan = dkpMan.FindPane(conPane_DetialList)
    If Not mblnShowDetial Then
        If Not objPan Is Nothing Then
            If Not objPan.Closed Then objPan.Close
        End If
        If rptData.Visible Then rptData.SetFocus
        Exit Sub
    Else
       If Not objPan Is Nothing Then
            If Not objPan.Selected Then
                objPan.Select
            End If
        End If
    End If
    Screen.MousePointer = vbHourglass
    Set obj�����¼�� = GetClinicRecordFromSignalSource(lng��ԴId)
    Call CPDPages.LoadData(obj�����¼��, Nothing, mobj���к�����λ, True)
    CPDPages.EditMode = ED_RegistPlan_View
    If rptData.Visible Then rptData.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
errHandle:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
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
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "����ΪXML�ļ�(&L)��", cbrControl.Index + 1)
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "���Ӻ�Դ(&J)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸ĺ�Դ(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ����Դ(&U)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "���ú�Դ(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "ͣ�ú�Դ(&T)")
    End With

    '�鿴�˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) 'ˢ����ǰ(���ʱע�ⷴ��)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowNoSourceDetial, "��ʾ��Դ������Ϣ(&C)", cbrControl.index)
        cbrControl.Checked = mblnShowDetial
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowStoped, "��ʾ��ͣ�ú�Դ(&S)", cbrControl.index)
        cbrControl.Checked = mblnShowStopSignal
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "���Ӻ�Դ", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸ĺ�Դ", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ����Դ", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "���ú�Դ", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "ͣ�ú�Դ", cbrControl.index + 1)
        .Item(cbrControl.index + 1).BeginGroup = True
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
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("B"), conMenu_Edit_Modify
        .Add 0, VK_DELETE, conMenu_Edit_Delete
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
    Dim blnVisible As Boolean, blnEnable As Boolean
    Dim blnStop As Boolean '�Ƿ���ͣ��
    
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    If rptData.SelectedRows.Count > 0 Then
        If Not rptData.SelectedRows(0).GroupRow Then
            blnEnable = rptData.SelectedRows(0).Record(COL_�Ƿ�ɾ��).Value = ""
            blnStop = blnEnable And rptData.SelectedRows(0).Record(COL_�Ƿ�ͣ��).Value <> ""
        End If
    End If
    blnVisible = zlStr.IsHavePrivs(mstrPrivs, "�����Դ����")

    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = rptData.Rows.Count > 0
    Case conMenu_EditPopup
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_NewItem
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And blnEnable And Not blnStop
    Case conMenu_Edit_Delete
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And blnEnable And Not blnStop
    Case conMenu_Edit_Reuse
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And blnStop
    Case conMenu_View_ShowNoSourceDetial
        Control.Checked = Control.Visible And mblnShowDetial
    Case conMenu_Edit_Stop
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And Not blnStop And blnEnable
    Case conMenu_View_FindType '���ҷ�ʽ
        Control.Caption = "��" & Decode(mintFindType, 0, "����", 1, "����", 2, "ҽ��", "����") & "���ˡ�"
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

Private Function ExcuteDelete() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ɾ������
    '���:lngID-��ԴID
    '����:���˺�
    '����:2016-03-30 14:37:59
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL  As String, rsTemp As ADODB.Recordset
    Dim lngID As Long, str���� As String
    On Error GoTo errHandle
    
    If rptData.SelectedRows.Count <= 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function

    lngID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
    str���� = Trim(rptData.SelectedRows(0).Record(COL_����).Caption)
    
    If Trim(rptData.SelectedRows(0).Record(COL_�Ƿ�ͣ��).Value) <> "" Then
        Call MsgBox("��Ҫɾ���ĺ���Ϊ" & str���� & "�ĺ�Դ�Ѿ���ͣ�ã�������ɾ����", vbInformation + vbOKOnly, gstrSysName)
        Exit Function
    End If
    If lngID = 0 Then
        MsgBox "��ǰδѡ��Ҫɾ���ĺ�Դ�����ܽ���ɾ��������", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If

    'ɾ����Ч�Լ��
    If CheckIsUserPreRegist(str����) Then
        If MsgBox("��ǰ��Դ(����Ϊ " & str���� & " )����ԤԼ�Һż�¼��ɾ���󣬽���Ըú�Դ�����г��ﰲ�Ž���ͣ��Ƿ����ɾ����", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        strSQL = "Select 1 From �ٴ������¼ Where ��ԴID=[1] And ��������+0>=Trunc(sysdate) And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
        If Not rsTemp.EOF Then
            If MsgBox("��ǰ��Դ(����Ϊ " & str���� & " )������Ч���ﰲ�ţ�ɾ���󣬽���Ըú�Դ����Щ���ﰲ�Ž���ͣ��Ƿ����ɾ����", _
                      vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Else
            If MsgBox("��ȷ��Ҫɾ����ǰ��Դ(����Ϊ " & str���� & " )��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    
    strSQL = "Zl_�ٴ������Դ_Delete(" & lngID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    'rptData.Records (rptData.SelectedRows(0).Index)
    ExcuteDelete = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim frmEdit As New frmClinicSignalSourceEdit, lngID As Long
    Dim str���� As String
    Err = 0: On Error GoTo errHandler
    If rptData.SelectedRows.Count > 0 Then
        If Not rptData.SelectedRows(0).GroupRow Then
            lngID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
            str���� = Trim(rptData.SelectedRows(0).Record(COL_����).Caption)
        End If
    End If
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_NewItem
        Dim strNewItem As String
        If frmEdit.ShowMe(Me, mlngModule, mstrPrivs, Fun_Add, , strNewItem) Then Call LoadData(, strNewItem)
    Case conMenu_Edit_Modify
        If frmEdit.ShowMe(Me, mlngModule, mstrPrivs, Fun_Update, lngID) Then Call LoadData
    Case conMenu_Edit_Delete
       If ExcuteDelete() Then Call LoadData
    Case conMenu_Edit_Reuse
        If StopAndResume(False) Then Call LoadData
    Case conMenu_Edit_Stop
        If StopAndResume(True) Then Call LoadData
    Case conMenu_View_ShowStoped '��ʾ��ͣ�ú�
        Control.Checked = Not Control.Checked
        mblnShowStopSignal = Control.Checked
        Call zlDatabase.SetPara("��ʾͣ�ú�Դ", IIf(mblnShowStopSignal, "1", "0"), glngSys, mlngModule)
        Call LoadData
    Case conMenu_View_ShowNoSourceDetial '��ʾ��ϸ��Ϣ
        mblnShowDetial = Not mblnShowDetial
        Control.Checked = mblnShowDetial
        Call zlDatabase.SetPara("��ʾȱʡ������Ϣ", IIf(mblnShowDetial, "1", "0"), glngSys, mlngModule)
        lngID = 0
        If rptData.SelectedRows.Count <> 0 Then
             If rptData.SelectedRows(0).GroupRow = False Then
                lngID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
             End If
        End If
        LoadDetialData (lngID)
        mlngPreSel��ԴID = lngID
    Case conMenu_View_Refresh
        Call GetRecords: Call ExecuteFilter
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
    
    If Not mrs��Դ Is Nothing Then
        With mrs��Դ
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
    Call LoadData(False)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitGrid()
    Dim i As Long
    Dim objCol As ReportColumn, lngIdx As Long
    
    Err = 0: On Error GoTo errHandler
    With rptData
        .AutoColumnSizing = False '��ʹ���Զ��п�
        .AllowColumnRemove = False '�������϶�ɾ����
        .ShowGroupBox = True '��ʾ�����
        .ShowItemsInGroups = False '����ʾ�ѷ������
        .MultipleSelection = False '���������ѡ��
        .SetImageList Me.imgList16
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid '�������߸�ʽ
            .HorizontalGridStyle = xtpGridSolid '�������߸�ʽ
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ������..."
            .ShadeSortColor = .BackColor
            Set .CaptionFont = Me.Font
            Set .TextFont = Me.Font
        End With
    End With

    With rptData.Columns
        Set objCol = .Add(COL_ID, "ID", 50, True): objCol.Visible = False
        Set objCol = .Add(COL_�Ű෽ʽͼ��, "", 20, False)
        objCol.Groupable = False
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.AllowRemove = False
        
        Set objCol = .Add(COL_����, "����", 50, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_����, "����", 50, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_����, "����", 100, True)
        Set objCol = .Add(COL_�շ���Ŀ, "�շ���Ŀ", 120, True)
        Set objCol = .Add(COL_ҽ��, "ҽ��", 80, True)
        Set objCol = .Add(COL_����, "����", 50, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_�Ű෽ʽ, "�Ű෽ʽ", 55, True)
        Set objCol = .Add(COL_ԤԼ����, "ԤԼ����", 55, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_����Ƶ��, "����Ƶ��", 55, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_���ջ���, "���ջ���", 55, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_���տ���״̬, "���տ���״̬", 100, True)
        Set objCol = .Add(COL_�ٴ��Ű�, "�ٴ��Ű�", 55, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_�����Ա�, "�����Ա�", 55, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_���������, "���������", 70, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_�Ƿ�ͣ��, "�Ƿ�ͣ��", 50, True): objCol.Visible = False
        Set objCol = .Add(COL_�Ƿ�ɾ��, "�Ƿ�ɾ��", 50, True): objCol.Visible = False
        Set objCol = .Add(COL_����ʱ��, "����ʱ��", 130, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_����ʱ��, "����ʱ��", 130, True): objCol.Alignment = xtpAlignmentCenter
    End With
    With rptData
    '        '������ȱʡ��������
        .SortOrder.DeleteAll
        .SortOrder.Add .Columns(COL_����)
        .SortOrder(0).SortAscending = True
        
        '�������������������
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns(COL_�Ű෽ʽ)
'        .GroupsOrder(0).SortAscending = True
        .Columns(COL_�Ű෽ʽ).Visible = False
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub GetRecords()
    '��ȡ��¼
    Dim strWhere As String, strSQL As String
    
    Err = 0: On Error GoTo errHandler
    Set mobj���к�����λ = GetUnitsObjects(GetUnitAll())
    
'    If mblnShowDeleteSignal = False Then '����ʾ��ɾ��
        strWhere = " And Nvl(a.�Ƿ�ɾ��,0) = 0"
'    End If
    If mblnShowStopSignal = False Then '����ʾ��ͣ��
        strWhere = strWhere & _
            " And Nvl(a.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate" & vbNewLine & _
            " And Nvl(b.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate" & vbNewLine & _
            " And Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate" & vbNewLine & _
            " And Nvl(d.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate"
    End If
    
    'û��"���п���"Ȩ�޵Ĳ���Աֻ�ܲ����Լ��������ҵĺ�Դ
    If HavePrivs(mstrPrivs, "���п���") = False Then
        strWhere = strWhere & "      And Exists (Select 1 From ������Ա Where ����id = a.����id And ��Աid = [1])"
    End If
    
    strSQL = "Select a.Id, a.����, a.����, c.���� As ����, c.���� As ���Ҽ���, b.���� As �շ���Ŀ," & vbNewLine & _
            "        a.ҽ������, d.���� As ҽ������,e.��ʶ��, a.ԤԼ����, a.����Ƶ��," & vbNewLine & _
            "        Nvl(a.�Ƿ񽨲���, 0) As �Ƿ񽨲���,nvl(a.�Ƿ��ٴ��Ű�,0) as �Ƿ��ٴ��Ű�," & vbNewLine & _
            "        Decode(nvl(a.���տ���״̬,0), 1, '����ԤԼ', 2, '��ֹԤԼ',3, '�ܽڼ������ÿ���', '���ϰ�') As ���տ���״̬," & vbNewLine & _
            "        Decode(nvl(a.�Ű෽ʽ,0), 1, '�����Ű�', 2, '�����Ű�', '�̶��Ű�') As �Ű෽ʽ," & vbNewLine & _
            "        Nvl(a.�Ƿ���ջ���, 0) As �Ƿ���ջ���, Nvl(a.�Ƿ�ɾ��, 0) As �Ƿ�ɾ��," & vbNewLine & _
            "        Case When Nvl(a.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate " & vbNewLine & _
            "               Or Nvl(b.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate " & vbNewLine & _
            "               Or Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate " & vbNewLine & _
            "               Or Nvl(d.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate Then 1 Else 0 End As �Ƿ�ͣ��," & vbNewLine & _
            "        a.����ʱ��, a.����ʱ��, a.�����Ա�, a.���������" & vbNewLine & _
            " From �ٴ������Դ A, �շ���ĿĿ¼ B, ���ű� C, ��Ա�� D,רҵ����ְ�� E" & vbNewLine & _
            " Where a.��Ŀid+0 = b.Id And a.����id = c.Id(+) And a.ҽ��ID = d.ID(+) and d.רҵ����ְ��=e.����(+)" & vbNewLine & _
            "        And Nvl(Nvl(c.վ��,[3]),Nvl([2],'-')) = Nvl([2],'-')" & vbNewLine & _
                strWhere
    Set mrs��Դ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, gstrNodeNo, gVisitPlan_ModulePara.str��Դά��վ��)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function LoadData(Optional ByVal blnReRead As Boolean = True, _
    Optional ByVal strNewItem As String) As Boolean
    '��������
    '��Σ�
    '   blnReRead �Ƿ����¶�ȡ����
    '   strNewItem ������Դ���룬���ڶ�λ
    Dim i As Long, j As Long, lngSelectRow As Long
    
    Err = 0: On Error GoTo errHandler
    Screen.MousePointer = vbHourglass

    If rptData.SelectedRows.Count > 0 Then lngSelectRow = rptData.SelectedRows(0).index
    rptData.Records.DeleteAll
    
    If mrs��Դ Is Nothing Then
        Call GetRecords
    ElseIf mrs��Դ.State <> adStateOpen Then
        Call GetRecords
    ElseIf blnReRead Then
        Call GetRecords
    End If
    
    Do While Not mrs��Դ.EOF
        Call InsertRowData(Nvl(mrs��Դ!ID), Nvl(mrs��Դ!����), Nvl(mrs��Դ!����), Nvl(mrs��Դ!����), _
            Nvl(mrs��Դ!�շ���Ŀ), Nvl(mrs��Դ!ҽ������), Nvl(mrs��Դ!��ʶ��), Nvl(mrs��Դ!�Ƿ񽨲���), Nvl(mrs��Դ!�Ű෽ʽ), _
            Nvl(mrs��Դ!ԤԼ����), Nvl(mrs��Դ!�����Ա�), Nvl(mrs��Դ!���������), _
            Nvl(mrs��Դ!����Ƶ��), Nvl(mrs��Դ!�Ƿ���ջ���), Nvl(mrs��Դ!���տ���״̬), Nvl(mrs��Դ!�Ƿ��ٴ��Ű�), _
            Nvl(mrs��Դ!�Ƿ�ͣ��), Nvl(mrs��Դ!�Ƿ�ɾ��), Nvl(mrs��Դ!����ʱ��), Nvl(mrs��Դ!����ʱ��))
        mrs��Դ.MoveNext
    Loop

    With rptData
        For i = 0 To .Records.Count - 1
            If i > .Records.Count - 1 Then Exit For
            If .Records(i).Item(COL_�Ƿ�ͣ��).Value <> "" _
                Or .Records(i).Item(COL_�Ƿ�ɾ��).Value <> "" Then
                For j = 0 To .Columns.Count - 1
                    .Records(i).Item(j).ForeColor = vbRed
                Next
            End If
        Next
    End With
    Call rptData.Populate '���������Ը��½���
    If rptData.Rows.Count > 0 Then '����ѡ������ʾ�ڿɼ�����
        If strNewItem <> "" Then
            For i = 0 To rptData.Rows.Count - 1
                If Not rptData.Rows(i).GroupRow Then
                    If rptData.Rows(i).Record(COL_����).Caption = strNewItem Then
                        rptData.FocusedRow = rptData.Rows(i)
                        Exit For
                    End If
                End If
            Next
        Else
            If lngSelectRow = 0 Then
                rptData.FocusedRow = rptData.Rows(0)
            ElseIf lngSelectRow > rptData.Rows.Count - 1 Then
                rptData.FocusedRow = rptData.Rows(rptData.Rows.Count - 1)
            Else
                rptData.FocusedRow = rptData.Rows(lngSelectRow)
            End If
        End If
    End If
    Call SetReportControlBackColorAlternate(rptData)
    
    mlngPreSel��ԴID = 0
    If rptData.SelectedRows.Count > 0 Then
         If rptData.SelectedRows(0).GroupRow = False Then
            mlngPreSel��ԴID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
         End If
    End If
    Call LoadDetialData(mlngPreSel��ԴID)
    
    Call mfrmMain.StatusShowInfoChanged(2, "��ǰ����" & mrs��Դ.RecordCount & "����Դ��Ϣ")
    
    Screen.MousePointer = vbDefault
    Exit Function
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InsertRowData(ByVal strID As String, ByVal str���� As String, ByVal str���� As String, ByVal str���� As String, _
    ByVal str��Ŀ As String, ByVal strҽ������ As String, ByVal str��ʶ�� As String, ByVal str���� As String, ByVal str�Ű෽ʽ As String, _
    ByVal strԤԼ���� As String, ByVal str�����Ա� As String, ByVal str��������� As String, _
    ByVal str����Ƶ�� As String, ByVal str���ջ��� As String, ByVal str���տ���״̬ As String, _
    ByVal str�ٴ��Ű� As String, ByVal str�Ƿ�ͣ�� As String, ByVal str�Ƿ�ɾ�� As String, _
    ByVal str����ʱ�� As String, ByVal str����ʱ�� As String)
    Dim objRecord As ReportRecord, ObjItem As ReportRecordItem
    Dim strTemp As String
    
    Err = 0: On Error GoTo errHandler
    With rptData
        Set objRecord = .Records.Add()
        Set ObjItem = objRecord.AddItem(strID)
        Set ObjItem = objRecord.AddItem("")
        
        'ͼ������
        Select Case str�Ű෽ʽ
        Case "�����Ű�"
            ObjItem.Icon = IIf(Val(str�Ƿ�ͣ��) = 0, 2, 3)
        Case "�����Ű�"
            ObjItem.Icon = IIf(Val(str�Ƿ�ͣ��) = 0, 4, 5)
        Case Else '�̶��Ű�
            ObjItem.Icon = IIf(Val(str�Ƿ�ͣ��) = 0, 0, 1)
        End Select
        
        Set ObjItem = objRecord.AddItem(str����)
        Set ObjItem = objRecord.AddItem(str����)
        If gVisitPlan_ModulePara.byt����ȽϷ�ʽ = 1 Then
            ObjItem.Caption = str����
            If Len(str����) >= 5 Then '���������Ժ󣬺�����ǰ�����ֵ������
                ObjItem.Value = str����
            Else
                ObjItem.Value = String(5 - Len(str����), "0") & str����
            End If
        End If
        Set ObjItem = objRecord.AddItem(str����)
        
        Set ObjItem = objRecord.AddItem(str��Ŀ)
        Set ObjItem = objRecord.AddItem(strҽ������)
        ObjItem.Caption = str��ʶ�� & strҽ������
        Set ObjItem = objRecord.AddItem(IIf(Val(str����) = 0, "", "��"))
        Set ObjItem = objRecord.AddItem(str�Ű෽ʽ)
        
        Set ObjItem = objRecord.AddItem(strԤԼ����)
        Set ObjItem = objRecord.AddItem(Val(str����Ƶ��))
        Set ObjItem = objRecord.AddItem(IIf(Val(str���ջ���) = 0, "", "��"))
        Set ObjItem = objRecord.AddItem(str���տ���״̬)
        Set ObjItem = objRecord.AddItem(IIf(Val(str�ٴ��Ű�) = 0, "", "��"))
        
        Set ObjItem = objRecord.AddItem(str�����Ա�)
        strTemp = str���������
        If InStr(strTemp, "~") > 0 Then
            If Split(strTemp, "~")(0) = "" Then
                strTemp = Split(strTemp, "~")(1) & "����"
            ElseIf Split(strTemp, "~")(1) = "" Then
                strTemp = Split(strTemp, "~")(0) & "����"
            End If
        End If
        Set ObjItem = objRecord.AddItem(strTemp)
        
        Set ObjItem = objRecord.AddItem(IIf(Val(str�Ƿ�ͣ��) = 0, "", "��"))
        Set ObjItem = objRecord.AddItem(IIf(Val(str�Ƿ�ɾ��) = 0, "", "��"))
        Set ObjItem = objRecord.AddItem(Format(str����ʱ��, "yyyy-mm-dd hh:mm:ss"))
        Set ObjItem = objRecord.AddItem(Format(str����ʱ��, "yyyy-mm-dd hh:mm:ss"))
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If Me.ActiveControl Is Nothing Then
        sccTitle.SetFocus
    ElseIf Not Me.ActiveControl Is txtFind(1) Then
        rptData.SetFocus
    End If
    Call mfrmMain.ActiveFormChange(Me)
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo errHandler
    mlngPreSel��ԴID = -1
    
    '��ȡ����ֵ
    mblnShowDetial = Val(zlDatabase.GetPara("��ʾȱʡ������Ϣ", glngSys, mlngModule, "0")) = 1
    mblnShowStopSignal = Val(zlDatabase.GetPara("��ʾͣ�ú�Դ", glngSys, mlngModule, "0")) = 1
    Call InitPancel
    Call InitGrid
    
    RestoreWinState Me, App.ProductName
    Dim strFindType As String
    Call GetRegInFor(g˽��ģ��, Me.Name, "FindType", strFindType)
    mintFindType = Val(strFindType)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    Call SaveRegInFor(g˽��ģ��, Me.Name, "FindType", mintFindType)
    If Not mrs��Դ Is Nothing Then Set mrs��Դ = Nothing
    If Not mrsWorkTime Is Nothing Then Set mrsWorkTime = Nothing
End Sub

Private Sub picBack_Resize()
    Err = 0: On Error Resume Next
    
    With picBack
        shpBorder.Move 0, 0, .ScaleWidth - 6, .ScaleHeight - 6
        sccTitle.Move .ScaleLeft, .ScaleTop, .ScaleWidth
        rptData.Left = .ScaleLeft + 10
        rptData.Top = sccTitle.Top + sccTitle.Height
        rptData.Width = .ScaleWidth - 30
        rptData.Height = .ScaleHeight - sccTitle.Height - 30
    End With
End Sub
 
Private Sub picDetailList_Resize()
    Err = 0: On Error Resume Next
    With picDetailList
        CPDPages.Left = .ScaleLeft
        CPDPages.Top = .ScaleTop
        CPDPages.Width = .ScaleWidth
        CPDPages.Height = .ScaleHeight
    End With
End Sub

Private Sub rptData_ColumnOrderChanged()
    Call SetReportControlBackColorAlternate(rptData)
End Sub

Private Sub rptData_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBarPopup
    
    Err = 0: On Error GoTo errHandler
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    Me.SetFocus: Call mfrmMain.ActiveFormChange(Me)
    
    Set objPopup = mcbsMain.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If objPopup Is Nothing Then Exit Sub
    If objPopup.Visible = False Then Exit Sub
    objPopup.CommandBar.ShowPopup
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub rptData_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim frmEdit As New frmClinicSignalSourceEdit, lngID As Long
    Dim blnStop As Boolean
    
    Err = 0: On Error GoTo errHandler
    If rptData.SelectedRows.Count = 0 Then Exit Sub
    If rptData.SelectedRows(0).GroupRow Then Exit Sub
    
    lngID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
    blnStop = rptData.SelectedRows(0).Record(COL_�Ƿ�ͣ��).Value <> ""
    If zlStr.IsHavePrivs(mstrPrivs, "�����Դ����") And blnStop = False Then
        If frmEdit.ShowMe(Me, mlngModule, mstrPrivs, Fun_Update, lngID) Then Call LoadData 'ˢ������
    Else
        frmEdit.ShowMe Me, mlngModule, mstrPrivs, Fun_View, lngID
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub rptData_SelectionChanged()
    Dim lng��ԴId As String
    
    Err = 0: On Error GoTo errHandler
    lng��ԴId = 0
    If rptData.SelectedRows.Count <> 0 Then
        With rptData.SelectedRows(0)
            If Not .GroupRow Then
                lng��ԴId = Val(.Record(COL_ID).Value)
            End If
        End With
    End If
    If mlngPreSel��ԴID = lng��ԴId Then Exit Sub
    
    mlngPreSel��ԴID = lng��ԴId
    Call LoadDetialData(lng��ԴId)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub rptData_SortOrderChanged()
    Call SetReportControlBackColorAlternate(rptData)
End Sub

Private Function StopAndResume(ByVal blnStop As Boolean) As Boolean
    '���ܣ�ͣ�û����ú�Դ
    '���أ�ͣ�û����óɹ�,����true,���򷵻�False
    Dim i As Integer, intRow As Integer
    Dim strSQL As String, str���� As String, lng��ԴId As Long
    Dim rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    If rptData.SelectedRows.Count = 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function
    
    str���� = Trim(rptData.SelectedRows(0).Record(COL_����).Caption)
    lng��ԴId = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
    
    If blnStop Then
        If CheckIsUserPreRegist(str����) Then
            If MsgBox("����Ϊ" & str���� & "�ĺ�Դ�Ѿ�����ԤԼ�Һż�¼��ͣ�øú�Դ�󣬽���Ըú�Դ�����г��ﰲ�Ž���ͣ��Ƿ����ͣ�ã�", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Else
            If MsgBox("ͣ�øú�Դ�󣬽���Ըú�Դ�����г��ﰲ�Ž���ͣ��Ƿ����ͣ�ú���Ϊ" & str���� & "�ĺ�Դ��", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    Else
        If MsgBox("��ȷ��Ҫ" & IIf(blnStop, "ͣ��", "����") & "����Ϊ""" & str���� & """�ĺ�Դ��", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        strSQL = "Select Case When Nvl(b.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate Then 1 Else 0 End As ����ͣ��," & vbNewLine & _
                "        Case When Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate Then 1 Else 0 End As ��Աͣ��," & vbNewLine & _
                "        Case When Nvl(d.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate Then 1 Else 0 End As ��Ŀͣ��" & vbNewLine & _
                " From �ٴ������Դ A, ���ű� B, ��Ա�� C, �շ���ĿĿ¼ D" & vbNewLine & _
                " Where a.����id = b.Id And a.ҽ��id = c.Id(+) And a.��ĿID = d.ID And a.Id = [1]" & vbNewLine & _
                "       And (Nvl(b.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate" & vbNewLine & _
                "            Or Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate" & vbNewLine & _
                "            Or Nvl(d.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate)"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ԴId)
        If Not rsTemp.EOF Then
            If Val(Nvl(rsTemp!��Աͣ��)) = 1 Then
                MsgBox "�ú�Դ��ҽ���ѱ�ͣ�û�ɾ������δ��������ҽ��ǰ�������øú�Դ��", vbInformation, gstrSysName
                Exit Function
            ElseIf Val(Nvl(rsTemp!����ͣ��)) = 1 Then
                MsgBox "�ú�Դ�Ŀ����ѱ�ͣ�û�ɾ������δ�������ÿ���ǰ�������øú�Դ��", vbInformation, gstrSysName
                Exit Function
            ElseIf Val(Nvl(rsTemp!��Ŀͣ��)) = 1 Then
                MsgBox "�ú�Դ���շ���Ŀ�ѱ�ͣ�û�ɾ������δ���������շ���Ŀǰ�������øú�Դ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    strSQL = "Zl_�ٴ������Դ_Stopandstart(" & lng��ԴId & "," & IIf(blnStop, 1, 0) & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    StopAndResume = True
    
    If blnStop = False Then '����ʱ��������û�����ɵĳ����¼
        '�������ɳ����¼
        strSQL = "Zl1_Auto_Buildingregisterplan(Null)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckIsUserPreRegist(ByVal str���� As String) As Boolean
    '����:����Ƿ����ԤԼ�Һ�
    '����:���ڷ���true,���򷵻�False
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    strSQL = "Select 1 From ������ü�¼" & vbNewLine & _
            " Where ��¼����=4 And ��¼״̬=0 And ����ʱ��>=Sysdate And ���㵥λ=[1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����)
    If Not rsTemp.EOF Then
        CheckIsUserPreRegist = True
        Exit Function
    End If

    strSQL = "Select 1 From ���˹Һż�¼" & vbNewLine & _
            " Where ��¼����=1 And ��¼״̬=1 And ����ʱ��>=Sysdate And �ű�=[1] And Rownum < 2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����)
    If Not rsTemp.EOF Then
        CheckIsUserPreRegist = True
        Exit Function
    End If
    CheckIsUserPreRegist = False
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub zlDataPrint(bytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If UserInfo.���� = "" Then Call GetUserInfo
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte, strHiddenCols As String
    
    Err = 0: On Error GoTo errHandler
    objOut.Title.Text = "�ٴ���Դ�嵥"
    '��ReportControlת��ΪVSFlexGrid
    strHiddenCols = CStr(COL_ID) & "," & CStr(COL_�Ű෽ʽͼ��) & "," & _
        CStr(COL_�Ƿ�ɾ��) & "," & CStr(COL_�Ƿ�ͣ��)
    Set objOut.Body = GetVsfGridData(rptData, strHiddenCols)
    
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True

    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    If bytMode = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytMode
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub sccTitle_GotFocus()
    On Error Resume Next
    If rptData.Visible Then rptData.SetFocus
End Sub

Private Sub txtFind_KeyPress(index As Integer, KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        Call ExecuteFilter
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
