VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmSegmentList 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   4350
      Left            =   90
      TabIndex        =   0
      Top             =   465
      Width           =   3405
      _Version        =   589884
      _ExtentX        =   6006
      _ExtentY        =   7673
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin VB.PictureBox picView 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   2145
      Left            =   90
      ScaleHeight     =   2145
      ScaleWidth      =   3450
      TabIndex        =   1
      Top             =   4920
      Width           =   3450
      Begin XtremeReportControl.ReportControl rptView 
         Height          =   2025
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3240
         _Version        =   589884
         _ExtentX        =   5715
         _ExtentY        =   3572
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2820
      Top             =   165
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSegmentList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSegmentList.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSegmentList.frx":0B34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmSegmentList.frx":0ECE
      Left            =   645
      Top             =   75
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmSegmentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ͼ�� = 0: ID: ���: ����
End Enum
Const con_UnDefine = -999
Const conPane_View = 201

'---------------------------------
'�����¼�
Public Event RowDblClick(ByVal ROW As XtremeReportControl.IReportRow)   '˫��һ�л������ϰ��س�
Public Event ModifiedOrDeleted(Action As Integer)                       '�޸Ļ�ɾ��ʾ��ʱ

'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mlngDemoId As Long          '��ǰʾ��id

Private mfrmParent As Form          '������
Private mlngFileID As Long          '�����ļ�id
Private mlngPatient As Long         '����id���ڲ��˲����༭ʱ������ȷ������ʾ���Ƿ�����
Private mlngVisit As Long           '��ҳid��Һŵ�ID
Private mlngAdvice As Long          'ҽ��ID

Private mintPower As Integer        'ʾ������Ȩ��Χ
'    mintPower=con_UnDefine��δ����;
'    mintPower=-1�����߱�ʾ������Ȩ;
'    mintPower=0��ȫԺ����ʱ��ʾ���е�ʾ����Ҳ���Ը���;
'    mintPower=1�����ң���ʱ��ʾȫԺͨ��ʾ��(����id is null)�����ڿ��ҹ��л�������Ա˽�е�ʾ���������ܸ���ȫԺͨ��ʾ��;
'    mintPower=2�����ˣ���ʱ��ʾȫԺͨ��ʾ��(����id is null)�����ڿ���ͨ��ʾ��(��Աid is null)�͸���ʾ����������ʾ���ɸ���

'-----------------------------------------------------
'��ʱ����
'-----------------------------------------------------


Dim lngCount As Long

'-----------------------------------------------------
'����Ϊ�ⲿ��������
'-----------------------------------------------------
Public Function zlRefresh(ByVal frmParent As Form) As Long
    '���ܣ�����ָ���ļ���ˢ���б�
    '������
    If frmParent.Name <> "frmMain" Then zlRefresh = 0: Exit Function
    
    Err = 0: On Error Resume Next
    With frmParent.Document
        mlngFileID = .EPRFileInfo.ID
        mlngPatient = .EPRPatiRecInfo.����ID
        mlngVisit = .EPRPatiRecInfo.��ҳID
        mlngAdvice = .EPRPatiRecInfo.ҽ��id
    End With
    Set mfrmParent = frmParent
    
    Err = 0: On Error GoTo 0
    zlRefresh = zlSubRefList(mlngDemoId)
End Function

'-----------------------------------------------------
'����Ϊ�ڲ���������
'-----------------------------------------------------
Private Function zlGetPower() As Integer
    '���ܣ���õ�ǰ�û���ʾ�������Ȩ��
    '���أ�ʾ������Ȩ����ֵ
    If mintPower = con_UnDefine Then
        If InStr(1, gstrPrivsEpr, "ȫԺ��������") <> 0 Then
            mintPower = 0
        ElseIf InStr(1, gstrPrivsEpr, "���Ҳ�������") <> 0 Then
            mintPower = 1
        ElseIf InStr(1, gstrPrivsEpr, "���˲�������") <> 0 Then
            mintPower = 2
        Else
            mintPower = -1
        End If
    End If
    zlGetPower = mintPower
End Function

Private Function zlSubRefList(Optional lngID As Long) As Long
    '���ܣ�ˢ��װ���嵥������λ��ָ���ļ�¼��
Dim rsTemp As New ADODB.Recordset
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

    gstrSQL = "Select l.Id, l.���, l.����, l.����, l.ͨ�ü�" & vbNewLine & _
            "From ��������Ŀ¼ l, Table(Cast(f_Segment_Usable([1], [2], [3], [4]) As " & gstrDbOwner & ".t_Dic_Rowset)) u" & vbNewLine & _
            "Where l.�ļ�id = [1] And Nvl(l.����, 0) = [5] And l.Id = To_Number(u.����)"
    Select Case mintPower
    Case 0
    Case 1
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.ͨ�ü�, 0) = 0 Or" & vbNewLine & _
                "      L.ͨ�ü� In (1, 2) And" & vbNewLine & _
                "      L.����id In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User))"

    Case Else
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.ͨ�ü�, 0) = 0 Or" & vbNewLine & _
                "      L.ͨ�ü� = 1 And" & vbNewLine & _
                "      L.����id In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User) Or" & vbNewLine & _
                "      L.ͨ�ü� = 2 And L.��Աid In (Select U.��Աid From �ϻ���Ա�� U Where U.�û��� = User))"
    End Select
    
    gstrSQL = gstrSQL & " Order By L.ͨ�ü� Desc, Lpad(L.���,13,'0') "
    
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmSegmentList", mlngFileID, mlngPatient, mlngVisit, mlngAdvice, 1)
    
    Me.rptList.Records.DeleteAll
    With rsTemp
        Do While Not .EOF
            Set rptRcd = Me.rptList.Records.Add()
            Set rptItem = rptRcd.AddItem(CInt(Val("" & !ͨ�ü�))): rptItem.Icon = rptItem.Value
            rptRcd.AddItem CLng(!ID)
            rptRcd.AddItem CStr("" & !���)
            rptRcd.AddItem CStr("" & !����)
            .MoveNext
        Loop
    End With
    Me.rptList.Populate
    If lngID <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If Val(rptRow.Record(mCol.ID).Value) = lngID Then
                Set Me.rptList.FocusedRow = rptRow: Exit For
            End If
        Next
    End If
    If Me.rptList.Rows.Count > 0 And (Me.rptList.FocusedRow Is Nothing) Then
        Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
    End If
    Call rptList_SelectionChanged
    zlSubRefList = Me.rptList.Records.Count
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlSubRefList = Me.rptList.Records.Count
End Function

'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRetuId As Long, strTemp As String
    
    '------------------------------------
    Select Case Control.ID
    Case conMenu_Edit_Modify
        lngRetuId = frmEPRModelEdit.ShowMe(mfrmParent, False, CByte(mintPower), mlngFileID, mlngDemoId)
        If lngRetuId = 0 Then Exit Sub
        Call zlSubRefList(lngRetuId)
        RaiseEvent ModifiedOrDeleted(1)
    Case conMenu_Edit_Delete
        Err = 0: On Error GoTo errHand
        strTemp = "���ɾ����ʾ����" & vbCrLf & "����" & Me.rptList.FocusedRow.Record(mCol.����).Value
        If MsgBox(strTemp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "zl_��������Ŀ¼_delete(" & mlngDemoId & ")"
        zlDatabase.ExecuteProcedure gstrSQL, "�ʾ��б�"
        With Me.rptList
            mlngDemoId = 0: lngRetuId = .FocusedRow.Index
            If .Rows.Count > lngRetuId + 1 Then
                mlngDemoId = .Rows(lngRetuId + 1).Record(mCol.ID).Value
            ElseIf lngRetuId > 0 Then
                mlngDemoId = .Rows(lngRetuId - 1).Record(mCol.ID).Value
            End If
        End With
        Call zlSubRefList(mlngDemoId)
        RaiseEvent ModifiedOrDeleted(2)
    Case conMenu_Edit_Request
        If frmEPRModelRequest.ShowMe(Me, mlngDemoId, mintPower) = True Then Call zlSubRefList(mlngDemoId)
    Case conMenu_View_Refresh
        Call zlSubRefList(mlngDemoId)
    End Select
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Sub cbsThis_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Err = 0: On Error Resume Next
    With Me.rptList
        .Left = Left: .Width = Right
        .Top = Top: .Height = Bottom - Top
    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_Request
        Control.Visible = (mintPower >= 0)
        Control.Enabled = (mlngDemoId <> 0)
        If Control.Enabled Then Control.Enabled = (Me.rptList.FocusedRow.Record(mCol.ͼ��).Value >= mintPower)
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_View: Item.Handle = Me.picView.hwnd
    End Select
End Sub

Private Sub Form_Load()
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim rptCol As ReportColumn
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gmstrPrivs�仯�����¿�����Ч
    mintPower = con_UnDefine
    mintPower = zlGetPower
    mlngDemoId = 0
    
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
    Me.cbsThis.ActiveMenuBar.Title = "�˵�": Me.cbsThis.ActiveMenuBar.Visible = False
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "����(&Q)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&V)"): cbrControl.BeginGroup = True
    End With
    
    '-----------------------------------------------------
    '����ʾ����ʾͣ������
    Dim panThis As Pane
    Set panThis = dkpMan.CreatePane(conPane_View, 450, 150, DockBottomOf, Nothing)
    panThis.Title = "ʾ������"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = False
    
    '-----------------------------------------------------
    With Me.rptList
        Set rptCol = .Columns.Add(mCol.ͼ��, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���, "���", 50, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.����, "����", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
    
    With Me.rptView
        Set rptCol = .Columns.Add(0, "���", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .ShowHeader = False
        .PreviewMode = True
        With .PaintManager
            .NoItemsText = "û�п���ʾ������..."
            .SetPreviewIndent 18, 0, 8, 6
        End With
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mfrmParent = Nothing
    imgList.ListImages.Clear
    ImageList_Destroy imgList.hImageList
End Sub

Private Sub picView_Resize()
    Err = 0: On Error Resume Next
    With Me.rptView
        .Left = 0: .Width = Me.picView.ScaleWidth
        .Top = 0: .Height = Me.picView.ScaleHeight
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.rptList
        If .Visible = False Then Exit Sub
        If .FocusedRow Is Nothing Then Exit Sub
        If .FocusedRow.GroupRow Then Exit Sub
        Call rptList_RowDblClick(.FocusedRow, .FocusedRow.Record.Item(mCol.ID))
    End With
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
Dim cbrPopupBar As CommandBar
Dim cbrPopupItem As CommandBarControl
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
    
    If Button <> vbRightButton Then Exit Sub
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.FindControl(xtpControlPopup, conMenu_EditPopup)
    If cbrMenuBar Is Nothing Then Exit Sub
    If cbrMenuBar.Visible = False Then Exit Sub
    
    Set cbrPopupBar = Me.cbsThis.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_RowDblClick(ByVal ROW As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Me.rptList.FocusedRow Is Nothing Then
        mlngDemoId = 0
    ElseIf Me.rptList.FocusedRow.GroupRow = True Then
        mlngDemoId = 0
    Else
        mlngDemoId = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
    End If
    If mlngDemoId = 0 Then Exit Sub
    RaiseEvent RowDblClick(Me.rptList.FocusedRow)
End Sub

Private Sub rptList_SelectionChanged()
Dim rsTemp As New ADODB.Recordset
Dim rsView As New ADODB.Recordset, strVSql As String, strView As String
Dim rptRcd As ReportRecord
    
    If Me.rptList.FocusedRow Is Nothing Then
        mlngDemoId = 0
    ElseIf Me.rptList.FocusedRow.GroupRow = True Then
        mlngDemoId = 0
    Else
        mlngDemoId = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
    End If
    If Me.Visible = False Then Exit Sub

    'ˢ��ʾ������
    gstrSQL = "Select Id, �����ı� From ������������ Where �ļ�id = [1] And �������� = 1 Order By �������"
    strVSql = "Select Id, ��������, �����ı�, �Ƿ���, Ҫ������" & vbNewLine & _
            "From ������������" & vbNewLine & _
            "Where �ļ�id = [1] And ��id = [2]" & vbNewLine & _
            "Order By �������"
    
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmSegmentList", mlngDemoId)
    Me.rptView.Records.DeleteAll
    Do While Not rsTemp.EOF
        Set rptRcd = Me.rptView.Records.Add()
        rptRcd.AddItem CStr(rsTemp!�����ı� & ":")
        strView = ""
        Set rsView = zlDatabase.OpenSQLRecord(strVSql, "frmSegmentList", mlngDemoId, CLng(rsTemp!ID))
        Do While Not rsView.EOF
            Select Case rsView!��������
            Case 2: strView = strView & rsView!�����ı�
            Case 3, 5: strView = strView & vbCrLf & "��" & vbCrLf
            Case 4: strView = strView & "[" & IIf(Trim("" & rsView!�����ı�) = "", rsView!Ҫ������, rsView!�����ı�) & "]"
            Case 7: strView = strView & "<" & rsView!�����ı� & ">"
            End Select
            strView = strView & IIf(Val("" & rsView!�Ƿ���) = 1, vbCrLf, "")
            rsView.MoveNext
        Loop
        rptRcd.PreviewText = Trim(strView)
        rsTemp.MoveNext
    Loop
    Me.rptView.Populate
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

