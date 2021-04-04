VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmClinicOfficeManage 
   BorderStyle     =   0  'None
   Caption         =   "�������ҹ���"
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeReportControl.ReportControl rptData 
      Height          =   3105
      Left            =   570
      TabIndex        =   0
      Top             =   720
      Width           =   6015
      _Version        =   589884
      _ExtentX        =   10610
      _ExtentY        =   5477
      _StockProps     =   0
      ShowGroupBox    =   -1  'True
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   6120
      MaxLength       =   100
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000C&
      Height          =   735
      Left            =   150
      Top             =   300
      Width           =   405
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   510
      TabIndex        =   1
      Top             =   300
      Width           =   5895
      _Version        =   589884
      _ExtentX        =   10398
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "��������>������������"
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
Attribute VB_Name = "frmClinicOfficeManage"
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
    COL_վ��
    COL_����
    COL_����
    COL_����
    COL_����
    COL_λ��
    COL_��æ��־
End Enum
Private mintFindType As Integer
Private mrsDoctorOffice As ADODB.Recordset

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
    Dim objCustom As CommandBarControlCustom
    
    Err = 0: On Error GoTo ErrHandler
    
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
    
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.id = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "��������(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�����(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ������(&D)")
    End With

    '�鿴�˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) 'ˢ����ǰ(���ʱע�ⷴ��)
'        Set cbrControl = .Add(xtpControlButton, conMenu_View_Notify, "ˢ������(&B)", cbrControl.Index)
        cbrControl.BeginGroup = True
    End With
    
    '����������
    '-----------------------------------------------------
    Set cbrToolBar = mcbsMain(2)
    For Each cbrControl In cbrToolBar.Controls '�����ǰ������һ��Control
        If Val(Left(cbrControl.id, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.id, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "��������", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�����", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ������", cbrControl.Index + 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
    End With
    
    Set objPopup = cbrToolBar.Controls.Add(xtpControlButtonPopup, conMenu_View_FindType, "�����ҹ��ˡ�")
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
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnVisible As Boolean, blnEnable As Boolean
    
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    blnVisible = zlStr.IsHavePrivs(mstrPrivs, "������������")
    If rptData.SelectedRows.Count > 0 Then
        blnEnable = Not rptData.SelectedRows(0).GroupRow
    End If
      
    Select Case Control.id
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
        Control.Enabled = Control.Visible And blnEnable
    Case conMenu_Edit_Delete
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And blnEnable
    Case conMenu_View_FindType '���ҷ�ʽ
        Control.Caption = "��" & Decode(mintFindType, 0, "����", 1, "����", 2, "վ��", "����") & "���ˡ�"
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 9 '���ҷ�ʽ
        Control.Checked = Val(Right(Control.id, 2)) - 1 = mintFindType
    End Select
End Sub

Public Sub InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If CommandBar.Parent Is Nothing Then Exit Sub
        
    Select Case CommandBar.Parent.id
    Case conMenu_View_FindType
        With CommandBar.Controls
            If .Count = 0 Then '��̬�Ӳ˵�,��1λ
                .Add xtpControlButton, conMenu_View_FindType * 100# + 1, "����(&1)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 2, "����(&2)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 3, "վ��(&3)"
            End If
        End With
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim frm As New frmClinicOfficeEdit, lngID As Long
    
    Err = 0: On Error GoTo ErrHandler
    If rptData.SelectedRows.Count > 0 Then
        If Not rptData.SelectedRows(0).GroupRow Then
            lngID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
        End If
    End If
    Select Case Control.id
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_NewItem
        Dim strNewItem As String
        If frm.ShowMe(Me, Fun_Add, , strNewItem) Then Call LoadData(, strNewItem)
    Case conMenu_Edit_Modify
        If frm.ShowMe(Me, Fun_Update, lngID) Then Call LoadData
    Case conMenu_Edit_Delete
        If ExcuteDelete() Then Call LoadData
    Case conMenu_View_Refresh
        Call GetRecords: Call ExecuteFilter
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 3 '���ҷ�ʽ
        mintFindType = Val(Right(Control.id, 2)) - 1
        mcbsMain.RecalcLayout
        txtFind(1).Text = ""
        If txtFind(1).Visible And txtFind(1).Enabled Then txtFind(1).SetFocus
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ExecuteFilter()
    '��������
    Dim strKey As String
    
    Err = 0: On Error GoTo ErrHandler
    Call zlControl.TxtSelAll(txtFind(1))
    
    If Not mrsDoctorOffice Is Nothing Then
        With mrsDoctorOffice
            If Trim(txtFind(1).Text) = "" Then
                .Filter = ""
            Else
                strKey = Replace(gstrLike, "%", "*") & UCase(txtFind(1).Text) & "*"
                Select Case mintFindType
                Case 0   '����(����)
                    .Filter = "���� Like '" & strKey & "' Or ���Ҽ��� Like '" & strKey & "'"
                Case 1   '����(����)
                    .Filter = "���� Like '" & strKey & "' Or ���� Like '" & strKey & "'"
                Case 2   'վ��
                    If Trim(txtFind(1).Text) = "ȫԺ" Then
                        .Filter = "վ������=null"
                    Else
                        .Filter = "վ������ Like '" & strKey & "'"
                    End If
                Case Else
                    .Filter = ""
                End Select
            End If
        End With
    End If
    If mintFindType = 8 Then mintFindType = 0 '���
    Call LoadData(False)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ExcuteDelete() As Boolean
    '����:ִ��ɾ������
    Dim strSQL  As String, rsTemp As ADODB.Recordset
    Dim lngID As Long, str���� As String
    
    On Error GoTo ErrHandler
    If rptData.SelectedRows.Count <= 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function

    lngID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
    str���� = Trim(rptData.SelectedRows(0).Record(COL_����).Value)
    
    If MsgBox("��ȷ��Ҫɾ�� " & str���� & " ��", _
        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
    '��飬��ʹ�õĲ���ɾ��
    If CheckHaveUsed(lngID) Then
      MsgBox "��ǰ�����ѱ�ʹ�ã�����ɾ����", vbInformation, gstrSysName: Exit Function
    End If

    'Zl_��������_Delete(
    strSQL = "Zl_��������_Delete("
    'Id_In ��������.Id%Type
    strSQL = strSQL & "" & lngID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    ExcuteDelete = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckHaveUsed(ByVal lng����ID As Long) As Boolean
    '��鵱ǰ�����Ƿ��ѱ�ʹ��
    Dim strSQL As String, rs���� As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    '���ԭ�ϰ�ʱ���Ƿ�ʹ�ã���ʹ�õĲ����޸�վ�㡢���ࡢʱ���
    '����ɾ����ʹ�õķ�Χ������һ��,��ʹ�õ�ʱ��ֻҪ��һ�����ɣ���ͬվ�㣬��ͬ������ܻ��ж��ͬ����ʱ��Σ�
    '�ٴ������Դ����
    strSQL = "Select 1 From �ٴ������Դ���� Where ����id = [1] And Rownum < 2"
    Set rs���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If Not rs���� Is Nothing Then
        If Not rs����.EOF Then CheckHaveUsed = True: Exit Function
    End If
    
    '�ٴ���������(�̶�����ģ��)
    strSQL = "Select 1 From �ٴ��������� Where ����id = [1] And Rownum < 2"
    Set rs���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If Not rs���� Is Nothing Then
        If Not rs����.EOF Then CheckHaveUsed = True: Exit Function
    End If
    
    '�ٴ��������Ҽ�¼
    strSQL = "Select 1 From �ٴ��������Ҽ�¼ Where ����id = [1] And Rownum < 2"
    Set rs���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If Not rs���� Is Nothing Then
        If Not rs����.EOF Then CheckHaveUsed = True: Exit Function
    End If
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitGrid()
    Dim i As Long
    Dim objCol As ReportColumn, lngIdx As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objItem As Field
    
    Err = 0: On Error GoTo ErrHandler
    With rptData
        .AutoColumnSizing = False '��ʹ���Զ��п�
        .AllowColumnRemove = False '�������϶�ɾ��������
        .ShowGroupBox = True '��ʾ�����
        .ShowItemsInGroups = False '����ʾ�ѷ������
        .MultipleSelection = False '���������ѡ��
'        .SetImageList Me.img16
        
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
        Set objCol = .Add(COL_վ��, "վ��", 50, True)
        Set objCol = .Add(COL_����, "����", 100, True)
        Set objCol = .Add(COL_����, "����", 60, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_����, "����", 100, True)
        Set objCol = .Add(COL_����, "����", 60, True)
        Set objCol = .Add(COL_λ��, "λ��", 100, True)
        Set objCol = .Add(COL_��æ��־, "��æ״̬", 80, True): objCol.Alignment = xtpAlignmentCenter
        
        '��̬�����û���չ�ֶ�,113315
        lngIdx = COL_��æ��־ + 1
        strSQL = "Select * From �������� Where 1 = 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������ұ�ṹ")
        For Each objItem In rsTemp.Fields
            If InStr(",ID,����,����,����,λ��,ȱʡ��־,վ��,", "," & UCase(objItem.Name) & ",") = 0 Then
                Set objCol = .Add(lngIdx, objItem.Name, 100, True): lngIdx = lngIdx + 1
                If objItem.Name Like "�Ƿ�*" Or (objItem.Type = adNumeric And objItem.Precision = 1) Then
                    objCol.Alignment = xtpAlignmentCenter
                End If
            End If
        Next
    End With
    With rptData
        '��վ��Ϳ��ҷ���
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns(COL_վ��)
        .GroupsOrder.Add .Columns(COL_����)
        .Columns(COL_վ��).Visible = False
        .Columns(COL_����).Visible = False
        
        '��վ��+��������(����)
        .SortOrder.DeleteAll
        .SortOrder.Add .Columns(COL_վ��)
        .SortOrder.Add .Columns(COL_����)
        .SortOrder(0).SortAscending = True
        .SortOrder(1).SortAscending = True
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub GetRecords()
    '��ȡ��¼
    Dim strSQL As String
    
    Err = 0: On Error GoTo ErrHandler
    strSQL = _
        "Select d.���� As վ������, c.���� As ����, c.���� As ���Ҽ���, a.*" & vbNewLine & _
        " From �������� A, �����������ÿ��� B, ���ű� C, Zlnodelist D" & vbNewLine & _
        " Where a.Id = b.����id(+) And b.����id = c.Id(+) And a.վ�� = d.���(+)" & vbNewLine & _
        "       And (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))"
    Set mrsDoctorOffice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Exit Sub
ErrHandler:
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
    '   strNewItem �����������ƣ����ڶ�λ
    Dim i As Long, j As Long
    Dim lngSelectRow As Long
    Dim objRecord As ReportRecord, objItem As ReportRecordItem
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objField As Field
    
    Err = 0: On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    If rptData.SelectedRows.Count > 0 Then lngSelectRow = rptData.SelectedRows(0).Index
    rptData.Records.DeleteAll
    
    If mrsDoctorOffice Is Nothing Then
        Call GetRecords
    ElseIf mrsDoctorOffice.State <> adStateOpen Then
        Call GetRecords
    ElseIf blnReRead Then
        Call GetRecords
    End If
    
    Do While Not mrsDoctorOffice.EOF
        Set objRecord = rptData.Records.Add()
        With objRecord
            Set objItem = .AddItem(Val(Nvl(mrsDoctorOffice!id)))
            Set objItem = .AddItem(Nvl(mrsDoctorOffice!վ������, "ȫԺ"))
            Set objItem = .AddItem(Nvl(mrsDoctorOffice!����, "����"))
            Set objItem = .AddItem(Nvl(mrsDoctorOffice!����))
            Set objItem = .AddItem(Nvl(mrsDoctorOffice!����))
            Set objItem = .AddItem(Nvl(mrsDoctorOffice!����))
            Set objItem = .AddItem(Nvl(mrsDoctorOffice!λ��))
            Set objItem = .AddItem(IIf(Val(Nvl(mrsDoctorOffice!ȱʡ��־)) = 0, "��", "æ"))
            
            '��̬�����û���չ�ֶ�,113315
            For Each objField In mrsDoctorOffice.Fields
                If InStr(",վ������,����,���Ҽ���,ID,����,����,����,λ��,ȱʡ��־,վ��,", "," & UCase(objField.Name) & ",") = 0 Then
                    If objField.Name Like "�Ƿ�*" Or (objField.Type = adNumeric And objField.Precision = 1) Then
                        Set objItem = .AddItem(IIf(Nvl(objField.Value) = "1", "��", ""))
                    ElseIf objField.Type = adDate Or objField.Type = adDBTimeStamp _
                        Or objField.Type = adDBDate Or objField.Type = adDBTime Then
                        Set objItem = .AddItem(Format(Nvl(objField.Value), "yyyy-mm-dd"))
                    Else
                        Set objItem = .AddItem(Nvl(objField.Value))
                    End If
                End If
            Next
        End With
        
        mrsDoctorOffice.MoveNext
    Loop

    Call rptData.Populate '���������Ը��½���
    With rptData
        If .Rows.Count > 0 Then '����ѡ������ʾ�ڿɼ�����
            If strNewItem <> "" Then
                For i = 0 To rptData.Rows.Count - 1
                    If Not rptData.Rows(i).GroupRow Then
                        If rptData.Rows(i).Record(COL_����).Value = strNewItem Then
                            rptData.FocusedRow = rptData.Rows(i)
                            Exit For
                        End If
                    End If
                Next
            Else
                If lngSelectRow = 0 Then
                    .FocusedRow = .Rows(0)
                ElseIf lngSelectRow > .Rows.Count - 1 Then
                    .FocusedRow = .Rows(.Rows.Count - 1)
                Else
                    .FocusedRow = .Rows(lngSelectRow)
                End If
            End If
        End If
    End With
    
    Call SetReportControlBackColorAlternate(rptData)
    Screen.MousePointer = vbDefault
    LoadData = True
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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
    Err = 0: On Error GoTo ErrHandler
    Call InitGrid
    RestoreWinState Me, App.ProductName
    
    Dim strFindType As String
    Call GetRegInFor(g˽��ģ��, Me.Name, "FindType", strFindType)
    mintFindType = Val(strFindType)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 0, 0, Me.ScaleWidth
    With rptData
        .Left = 10: .Top = sccTitle.Top + sccTitle.Height
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - 10
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    Call SaveRegInFor(g˽��ģ��, Me.Name, "FindType", mintFindType)
    If Not mrsDoctorOffice Is Nothing Then Set mrsDoctorOffice = Nothing
End Sub

Private Sub rptData_ColumnOrderChanged()
    Call SetReportControlBackColorAlternate(rptData)
End Sub

Private Sub rptData_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBarPopup
    
    Err = 0: On Error GoTo ErrHandler
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    Me.SetFocus: Call mfrmMain.ActiveFormChange(Me)
    
    Set objPopup = mcbsMain.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If objPopup Is Nothing Then Exit Sub
    If objPopup.Visible = False Then Exit Sub
    objPopup.CommandBar.ShowPopup
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub rptData_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim frm As New frmClinicOfficeEdit, lngID As Long
    
    Err = 0: On Error GoTo ErrHandler
    If rptData.SelectedRows.Count > 0 Then
        If Not rptData.SelectedRows(0).GroupRow Then
            lngID = rptData.SelectedRows(0).Record(COL_ID).Value
            If zlStr.IsHavePrivs(mstrPrivs, "������������") Then
                If frm.ShowMe(Me, Fun_Update, lngID) Then Call LoadData 'ˢ������
            Else
                frm.ShowMe Me, Fun_View, lngID
            End If
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub rptData_SortOrderChanged()
    Call SetReportControlBackColorAlternate(rptData)
End Sub

Private Sub zlDataPrint(BytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If UserInfo.���� = "" Then Call GetUserInfo
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte
    
    Err = 0: On Error GoTo ErrHandler
    objOut.Title.Text = "���������嵥"
    '��ReportControlת��ΪVSFlexGrid
    Set objOut.Body = GetVsfGridData(rptData, CStr(COL_ID))

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
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub sccTitle_GotFocus()
    On Error Resume Next
    If rptData.Visible Then rptData.SetFocus
End Sub

Private Sub txtFind_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        Call ExecuteFilter
        If rptData.Visible Then rptData.SetFocus
    End If
End Sub

Private Sub txtFind_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 93 Then
        '�����Ҽ��˵���ݼ������ճ��������
        If Clipboard.GetText <> "" Then Clipboard.Clear
    End If
End Sub

Private Sub txtFind_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtFind(Index).Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtFind(Index).Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub
Private Sub txtFind_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtFind(Index).Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub
