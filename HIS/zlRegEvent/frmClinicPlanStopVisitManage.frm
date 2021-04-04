VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmClinicPlanStopVisitManage 
   Caption         =   "ͣ�ﰲ�Ź���"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9645
   Icon            =   "frmClinicPlanStopVisitManage.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   9645
   StartUpPosition =   2  '��Ļ����
   Begin XtremeReportControl.ReportControl rptData 
      Height          =   2505
      Left            =   1830
      TabIndex        =   0
      Top             =   1200
      Width           =   5385
      _Version        =   589884
      _ExtentX        =   9499
      _ExtentY        =   4419
      _StockProps     =   0
      ShowGroupBox    =   -1  'True
   End
   Begin VB.PictureBox picButton 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   9645
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   5940
      Visible         =   0   'False
      Width           =   9645
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   450
         TabIndex        =   4
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "�˳�(&E)"
         Height          =   350
         Left            =   7830
         TabIndex        =   3
         Top             =   180
         Width           =   1100
      End
   End
   Begin ComctlLib.ImageList imgList16 
      Left            =   6690
      Top             =   4500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicPlanStopVisitManage.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicPlanStopVisitManage.frx":0E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicPlanStopVisitManage.frx":13FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicPlanStopVisitManage.frx":1998
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   1350
      TabIndex        =   1
      Top             =   480
      Width           =   7905
      _Version        =   589884
      _ExtentX        =   13944
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "���ﰲ��>ͣ�ﰲ��"
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
      Height          =   495
      Left            =   240
      Top             =   180
      Width           =   915
   End
End
Attribute VB_Name = "frmClinicPlanStopVisitManage"
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
    Col_��¼ID
    COL_ͼ��
    Col_״̬
    COL_������
    COL_ͣ�����
    COL_��ʼʱ��
    COL_��ֹʱ��
    COL_ͣ��ԭ��
    COL_����ʱ��
    COL_������
    COL_����ʱ��
    COL_ʧЧʱ��
    COL_�Ǽ���
End Enum
Private mstrFilter As String
Private mstrDefaultFilter  As String
Private Type Type_SQLCondition
    ApplyName As String
    AuditName As String
    StopBegin As Date
    StopEnd As Date
End Type
Private SQLCondition As Type_SQLCondition

Private mstrDoctorName As String
Private mblnShowDoctorStopVisit As Boolean

Public Sub ShowDoctorStopVisit(frmParent As Form, ByVal strDoctorName As String)
    '��ʾָ��ҽ��ͣ����Ϣ
    mstrDoctorName = strDoctorName
    
    If strDoctorName = "" Then Exit Sub
    On Error Resume Next
    mblnShowDoctorStopVisit = True
    Me.Show 1, frmParent
End Sub

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, _
    ByVal strPrivs As String, ByVal lngModule As Long)
    '��ʼ������
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    
    mstrPrivs = strPrivs
    mlngModule = lngModule
End Sub

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddMonthPlan, "�ƶ��³����(&Y)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddWeekPlan, "�ƶ��ܳ����(&W)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "ͣ������(&A)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ȡ������(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "ͣ������(&V)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnAudit, "ȡ������(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "��ֹ����(&C)"): cbrControl.BeginGroup = True
    End With

    '�鿴�˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) 'ˢ����ǰ(���ʱע�ⷴ��)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "����(&F)", cbrControl.index)
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddMonthPlan, "�³����", cbrControl.index + 1): cbrControl.BeginGroup = True
        cbrControl.ToolTipText = "�ƶ��³����"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddWeekPlan, "�ܳ����", cbrControl.index + 1)
        cbrControl.ToolTipText = "�ƶ��ܳ����"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "ͣ������", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ȡ������", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "ͣ������", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnAudit, "ȡ������", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "��ֹ����", cbrControl.index + 1): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "����", cbrControl.index + 1): cbrControl.BeginGroup = True
    End With
    
    '����Ŀ����
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("Y"), conMenu_Edit_AddMonthPlan
        .Add FCONTROL, Asc("W"), conMenu_Edit_AddWeekPlan
        
        .Add FCONTROL, Asc("D"), conMenu_Edit_Delete
        .Add FCONTROL, Asc("V"), conMenu_Edit_Audit
        .Add FCONTROL, Asc("C"), conMenu_Edit_UnAudit
        .Add FCONTROL, Asc("F"), conMenu_View_Filter
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
    Dim blnEnable As Boolean, blnAudit As Boolean, blnStop As Boolean
    
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    If Control.ID = conMenu_Edit_Delete Or Control.ID = conMenu_Edit_Audit _
        Or Control.ID = conMenu_Edit_UnAudit Or Control.ID = conMenu_Edit_Stop Then
        If rptData.SelectedRows.Count > 0 Then
            If Not rptData.SelectedRows(0).GroupRow Then
                blnAudit = rptData.SelectedRows(0).Record(COL_������).Value <> ""
                blnEnable = Val(rptData.SelectedRows(0).Record(Col_��¼ID).Value) = 0 '�г����¼�Ĳ������κβ���
                blnStop = rptData.SelectedRows(0).Record(COL_ʧЧʱ��).Value <> ""
            End If
        End If
    End If

    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = rptData.Rows.Count > 0
    Case conMenu_EditPopup
        Control.Visible = (mfrmMain.mFunListActived And (HavePrivs(mstrPrivs, "���ﰲ��"))) _
            Or (mfrmMain.mFunListActived = False And (HavePrivs(mstrPrivs, "���ﰲ��;ͣ������;ͣ������")))
        Control.Enabled = Control.Visible
    Case conMenu_Edit_AddMonthPlan '�ƶ��³����
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And mfrmMain.mFunListActived
        Control.Enabled = Control.Visible
    Case conMenu_Edit_AddWeekPlan '�ƶ��ܳ����
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��") And mfrmMain.mFunListActived
        Control.Enabled = Control.Visible
    Case conMenu_Edit_NewItem 'ͣ������
        Control.Visible = HavePrivs(mstrPrivs, "ͣ������")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Delete 'ȡ������
        Control.Visible = HavePrivs(mstrPrivs, "ͣ������")
        Control.Enabled = Control.Visible And blnEnable And Not blnAudit
    Case conMenu_Edit_Audit 'ͣ������
        Control.Visible = HavePrivs(mstrPrivs, "ͣ������")
        Control.Enabled = Control.Visible And blnEnable And Not blnAudit
    Case conMenu_Edit_UnAudit 'ȡ������
        Control.Visible = HavePrivs(mstrPrivs, "ͣ������")
        Control.Enabled = Control.Visible And blnEnable And blnAudit And Not blnStop
    Case conMenu_Edit_Stop '��ֹ����
        Control.Visible = HavePrivs(mstrPrivs, "ͣ������")
        Control.Enabled = Control.Visible And blnEnable And blnAudit And Not blnStop
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim frmEdit As New frmClinicPlanStopVisitEdit, lngID As Long
    Dim strApplyName As String
    
    Err = 0: On Error GoTo errHandler
    If rptData.SelectedRows.Count > 0 Then
        If Not rptData.SelectedRows(0).GroupRow Then
            lngID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
        End If
    End If
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_NewItem 'ͣ������
        If frmEdit.ShowMe(Me, mlngModule, mstrPrivs, 1, , strApplyName) Then Call LoadData(strApplyName)
    Case conMenu_Edit_Delete 'ȡ������
        If frmEdit.ShowMe(Me, mlngModule, mstrPrivs, 2, lngID) Then Call LoadData
    Case conMenu_Edit_Audit 'ͣ������
        If frmEdit.ShowMe(Me, mlngModule, mstrPrivs, 3, lngID) Then Call LoadData
    Case conMenu_Edit_UnAudit 'ȡ������
        If frmEdit.ShowMe(Me, mlngModule, mstrPrivs, 4, lngID) Then Call LoadData
    Case conMenu_Edit_Stop '��ֹ����
        If frmEdit.ShowMe(Me, mlngModule, mstrPrivs, 5, lngID) Then Call LoadData
    Case conMenu_View_Refresh
        Call LoadData 'ˢ������
    Case conMenu_View_Filter '����
        With frmClinicPlanStopVisitFilter
            .mblnOk = False
            .Show 1, Me
            If .mblnOk Then
                mstrFilter = .mstrFilter
                SQLCondition.ApplyName = Trim(.txtApply.Text)
                SQLCondition.AuditName = Trim(.txtAudit.Text)
                SQLCondition.StopBegin = .dtpStopBegin.Value
                SQLCondition.StopEnd = .dtpStopEnd.Value
                Call LoadData
            End If
        End With
    End Select
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub RefreshData()
    Err = 0: On Error GoTo errHandler
    
    Call LoadData
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If mblnShowDoctorStopVisit Then Exit Sub
    If Me.ActiveControl Is Nothing Then
        sccTitle.SetFocus
    Else
        rptData.SetFocus
    End If
    Call mfrmMain.ActiveFormChange(Me)
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo errHandler
    
    Call InitGrid
    mstrFilter = " And Nvl(ʧЧʱ��,��ֹʱ��)>sysdate"
    
    If mblnShowDoctorStopVisit Then
        If LoadData() = False Then
            MsgBox "ҽ�� " & mstrDoctorName & " ��ǰ����Чͣ�ﰲ�ţ�", vbInformation + vbOKOnly, gstrSysName
            Unload Me: Exit Sub
        End If
        shpBorder.Visible = False
        sccTitle.Visible = False
        Me.Caption = mstrDoctorName & " ͣ�ﰲ��"
        
        picButton.Visible = True
    End If
    RestoreWinState Me, App.ProductName
    '�Ƿ���ʾ�����
    rptData.ShowGroupBox = (mblnShowDoctorStopVisit = False)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    If mblnShowDoctorStopVisit = False Then
        shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
        sccTitle.Move 8, 8, shpBorder.Width - 20
    End If
    
    With rptData
        .Left = 8
        .Top = IIf(mblnShowDoctorStopVisit, 0, sccTitle.Top + sccTitle.Height)
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - IIf(mblnShowDoctorStopVisit, picButton.Height, 0) - .Top - 20
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrFilter = ""
    mblnShowDoctorStopVisit = False
    SaveWinState Me, App.ProductName
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
            Set .PreviewTextFont = Me.Font
        End With
    End With
    
    With rptData.Columns
        Set objCol = .Add(COL_ID, "ID", 50, True): objCol.Visible = False
        Set objCol = .Add(Col_��¼ID, "��¼ID", 50, True): objCol.Visible = False
        Set objCol = .Add(COL_ͼ��, "", 20, False)
        objCol.Groupable = False
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.AllowRemove = False
        Set objCol = .Add(Col_״̬, "״̬", 50, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_������, "������", 50, True)
        Set objCol = .Add(COL_ͣ�����, "ͣ�����", 100, True)
        Set objCol = .Add(COL_��ʼʱ��, "��ʼʱ��", 130, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_��ֹʱ��, "��ֹʱ��", 130, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_ͣ��ԭ��, "ͣ��ԭ��", 140, True)
        Set objCol = .Add(COL_����ʱ��, "����ʱ��", 130, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_������, "������", 50, True)
        Set objCol = .Add(COL_����ʱ��, "����ʱ��", 130, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_ʧЧʱ��, "ʧЧʱ��", 130, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_�Ǽ���, "�Ǽ���", 50, True)
    End With
    
    With rptData
        '�������˷�����ʾ
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns(COL_������)
        .Columns(COL_������).Visible = False
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function LoadData(Optional ByVal strApplyName As String) As Boolean
    '��Σ�
    '   strApplyName - ȱʡ��λ����ҽ��
    Dim i As Long, j As Long
    Dim lngSelectRow As Long
    Dim strSQL As String, rsData As ADODB.Recordset
    Dim objRecord As ReportRecord, ObjItem As ReportRecordItem
    Dim dtNow As Date
    
    Err = 0: On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    If rptData.SelectedRows.Count > 0 Then lngSelectRow = rptData.SelectedRows(0).index
    rptData.Records.DeleteAll
    
    If mblnShowDoctorStopVisit Then mstrFilter = " And Nvl(ʧЧʱ��,��ֹʱ��)>sysdate And ������=[1]"
    strSQL = "Select ID, ��¼ID, ͣ��ԭ��, ��ʼʱ��, ��ֹʱ��, ������, ����ʱ��, ������, ����ʱ��," & vbNewLine & _
            "       Decode(����ʱ��,NULL,'δ����','������') As ״̬, �Ǽ���, ʧЧʱ��,Nvl(ͣ�����,'-') As ͣ�����" & vbNewLine & _
            " From �ٴ�����ͣ���¼" & vbNewLine & _
            " Where ��¼ID Is Null " & mstrFilter
    If mblnShowDoctorStopVisit = False Then
        With SQLCondition
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .ApplyName, .AuditName, .StopBegin, .StopEnd)
        End With
    Else
        strSQL = strSQL & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select ID, ��¼id, ͣ��ԭ��, ��ʼʱ��, ��ֹʱ��, ������, ����ʱ��, ������, ����ʱ��," & vbNewLine & _
            "        Decode(����ʱ��, Null, 'δ����', '������') As ״̬, �Ǽ���, ʧЧʱ��,Nvl(ͣ�����,'-') As ͣ�����" & vbNewLine & _
            " From �ٴ�����ͣ���¼ A" & vbNewLine & _
            " Where ��¼id Is Not Null " & mstrFilter & vbNewLine & _
            "       And Not Exists(Select 1" & vbNewLine & _
            "                       From �ٴ�����ͣ���¼" & vbNewLine & _
            "                       Where ��¼id Is Null " & mstrFilter & " And a.��ʼʱ�� >= ��ʼʱ�� And Nvl(a.ʧЧʱ��,a.��ֹʱ��) <= Nvl(ʧЧʱ��,��ֹʱ��))"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrDoctorName)
    End If
    If rsData.RecordCount = 0 Then
        Screen.MousePointer = vbDefault
        Call rptData.Populate '���������Ը��½���
        Exit Function
    End If
    
    dtNow = zlDatabase.Currentdate
    Do While Not rsData.EOF
        Set objRecord = rptData.Records.Add()
        objRecord.AddItem Nvl(rsData!ID)
        objRecord.AddItem Nvl(rsData!��¼ID)
        'ͼ������
        Set ObjItem = objRecord.AddItem("")
        If Nvl(rsData!����ʱ��) = "" Then 'δ����
            ObjItem.Icon = IIf(CDate(Nvl(rsData!ʧЧʱ��, rsData!��ֹʱ��)) > dtNow, 0, 1)
        Else '������
            ObjItem.Icon = IIf(CDate(Nvl(rsData!ʧЧʱ��, rsData!��ֹʱ��)) > dtNow, 2, 3)
        End If
        objRecord.AddItem Nvl(rsData!״̬)
        objRecord.AddItem Nvl(rsData!������)
        objRecord.AddItem Nvl(rsData!ͣ�����)
        objRecord.AddItem Format(Nvl(rsData!��ʼʱ��), "yyyy-mm-dd hh:mm:ss")
        objRecord.AddItem Format(Nvl(rsData!��ֹʱ��), "yyyy-mm-dd hh:mm:ss")
        objRecord.AddItem Nvl(rsData!ͣ��ԭ��)
        objRecord.AddItem Format(Nvl(rsData!����ʱ��), "yyyy-mm-dd hh:mm:ss")
        objRecord.AddItem Nvl(rsData!������)
        objRecord.AddItem Format(Nvl(rsData!����ʱ��), "yyyy-mm-dd hh:mm:ss")
        objRecord.AddItem Format(Nvl(rsData!ʧЧʱ��), "yyyy-mm-dd hh:mm:ss")
        objRecord.AddItem Format(Nvl(rsData!�Ǽ���), "yyyy-mm-dd hh:mm:ss")
        rsData.MoveNext
    Loop
    Call rptData.Populate '���������Ը��½���
    '����������ɫ��ʾ
    For i = 0 To rptData.Records.Count - 1
        If rptData.Records(i).Item(COL_����ʱ��).Value <> "" Then
            For j = 0 To rptData.Columns.Count - 1
                rptData.Records(i).Item(j).ForeColor = vbBlue
            Next
        End If
    Next
    
    If rptData.Rows.Count > 0 Then '����ѡ������ʾ�ڿɼ�����
        If strApplyName <> "" Then
            For i = 0 To rptData.Rows.Count - 1
                If Not rptData.Rows(i).GroupRow Then
                    If rptData.Rows(i).Record(COL_������).Value = strApplyName Then
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
    Screen.MousePointer = vbDefault
    LoadData = True
    Exit Function
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub picButton_Resize()
    On Error Resume Next
    cmdExit.Left = picButton.ScaleWidth - cmdExit.Width - 500
End Sub

Private Sub rptData_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBarPopup
    
    Err = 0: On Error GoTo errHandler
    If mblnShowDoctorStopVisit Then Exit Sub
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

Private Sub sccTitle_GotFocus()
    On Error Resume Next
    If rptData.Visible Then rptData.SetFocus
End Sub

Private Sub zlDataPrint(bytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If UserInfo.���� = "" Then Call GetUserInfo
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte, strHiddenCols As String
    
    Err = 0: On Error GoTo errHandler
    objOut.Title.Text = "ͣ�ﰲ���嵥"
    '��ReportControlת��ΪVSFlexGrid
    strHiddenCols = CStr(COL_ID) & "," & CStr(Col_��¼ID) & "," & CStr(COL_ͼ��)
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

