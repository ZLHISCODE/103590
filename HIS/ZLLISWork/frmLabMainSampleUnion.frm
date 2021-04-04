VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmLabMainSampleUnion 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   2775
      Left            =   570
      TabIndex        =   0
      Top             =   30
      Width           =   5745
      _Version        =   589884
      _ExtentX        =   10134
      _ExtentY        =   4895
      _StockProps     =   0
      BorderStyle     =   2
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      ShowItemsInGroups=   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl RptComPare1 
      Height          =   2775
      Left            =   450
      TabIndex        =   1
      Top             =   3090
      Width           =   3075
      _Version        =   589884
      _ExtentX        =   5424
      _ExtentY        =   4895
      _StockProps     =   0
      BorderStyle     =   2
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      ShowItemsInGroups=   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl RptComPare2 
      Height          =   2805
      Left            =   4350
      TabIndex        =   2
      Top             =   3090
      Width           =   3075
      _Version        =   589884
      _ExtentX        =   5424
      _ExtentY        =   4948
      _StockProps     =   0
      BorderStyle     =   2
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      ShowItemsInGroups=   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   6930
      Top             =   510
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   7080
      Top             =   2070
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmLabMainSampleUnion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event StartEdit(Cancel As Boolean)
Private Enum mCol
    �걾ID
    ����
    �Ա�
    ����
    ������Ŀ
    ��ʶ��
    ����
    �걾ʱ��
    ������
    �ϲ�״̬
End Enum
Private Enum mRCol
    ������Ŀ
    ���
    ��λ
    ��־
    �ο�
End Enum

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intLoop As Integer
    Dim intNext As Integer
    Dim lngID As Long
    
    Select Case Control.ID
        Case conMenu_Edit_Insert                        '���úϲ��걾
            If Not Me.rptList.FocusedRow Is Nothing Then
                lngID = Me.rptList.FocusedRow.Record.Item(mCol.�걾ID).Value
                For intLoop = 1 To Me.rptList.Records.Count
                    '����д��ϲ��򱻺ϲ��걾
                    If Me.rptList.Records(intLoop - 1).Item(mCol.�걾ID).Value = lngID Then
                        Me.rptList.Records(intLoop - 1).Item(mCol.�ϲ�״̬).Value = "�ϲ��걾"
                    Else
                        Me.rptList.Records(intLoop - 1).Item(mCol.�ϲ�״̬).Value = "���ϲ��걾"
                    End If
                    '������ɫ
                    For intNext = 0 To Me.rptList.Columns.Count
                        If Me.rptList.Records(intLoop - 1).Item(mCol.�ϲ�״̬).Value = "�ϲ��걾" Then
                            Me.rptList.Records(intLoop - 1).Item(intNext).ForeColor = vbBlue
                        Else
                            Me.rptList.Records(intLoop - 1).Item(intNext).ForeColor = vbRed
                        End If
                    Next
                Next
                Me.rptList.Populate
            End If
        Case conMenu_Manage_ThingDel                    '������кϲ��걾
            Me.rptList.Records.DeleteAll
            Me.rptList.Populate
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = Me.rptList.hWnd
        Case 2
            Item.Handle = Me.RptComPare1.hWnd
        Case 3
            Item.Handle = Me.RptComPare2.hWnd
    End Select
End Sub

Private Sub Form_Load()
    Dim Column As ReportColumn
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, Pane4 As Pane, Pane5 As Pane
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Set Me.cbrthis.Icons = zlCommFun.GetPubIcons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False

    dkpMain.Options.DefaultPaneOptions = PaneNoCloseable
    dkpMain.Options.HideClient = True
    
    Set Pane1 = dkpMain.CreatePane(1, 200, 50, DockTopOf, Nothing)
    Pane1.Title = "�ϲ��б�"
    Pane1.Handle = Me.rptList.hWnd
    Pane1.Options = PaneNoCaption

    Set Pane2 = dkpMain.CreatePane(2, 200, 150, DockBottomOf, Nothing)
    Pane2.Title = "�ϲ�����嵥"
    Pane2.Handle = Me.RptComPare1.hWnd
    Pane2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set Pane3 = dkpMain.CreatePane(3, 200, 150, DockRightOf, Pane2)
    Pane3.Title = "���ϲ�����嵥"
    Pane3.Handle = Me.RptComPare2.hWnd
    Pane3.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    With Me.rptList.Columns
        rptList.AllowColumnRemove = False
        rptList.ShowItemsInGroups = False
        
        With rptList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "˫����ߵ��б��еĲ������ӵ��ϲ��б���..."
            .VerticalGridStyle = xtpGridSolid
        End With
        
        Set Column = .Add(mCol.�걾ID, "�걾ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.����, "����", 75, True)
        Set Column = .Add(mCol.�Ա�, "�Ա�", 40, True)
        Set Column = .Add(mCol.����, "����", 40, True)
        Set Column = .Add(mCol.����, "����", 100, True)
        Set Column = .Add(mCol.��ʶ��, "��ʶ��", 80, True)
        Set Column = .Add(mCol.�걾ʱ��, "�걾ʱ��", 75, False)
        Set Column = .Add(mCol.������, "������", 75, True)
        Set Column = .Add(mCol.�ϲ�״̬, "�ϲ�״̬", 75, True)
    End With
    
    With Me.RptComPare1.Columns
        RptComPare1.AllowColumnRemove = False
        RptComPare1.ShowItemsInGroups = False
        
        With RptComPare1.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "�ϲ��ļ�����ĿΪ��..."
            .VerticalGridStyle = xtpGridSolid
        End With
        Set Column = .Add(mRCol.������Ŀ, "������Ŀ", 85, True)
        Set Column = .Add(mRCol.���, "���", 85, True)
        Set Column = .Add(mRCol.��λ, "��λ", 65, True)
        Set Column = .Add(mRCol.��־, "��־", 65, True)
        Set Column = .Add(mRCol.�ο�, "�ο�", 85, True)
    End With
    
    With Me.RptComPare2.Columns
        RptComPare2.AllowColumnRemove = False
        RptComPare2.ShowItemsInGroups = False
        
        With RptComPare2.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "���ϲ��ļ�����ĿΪ��..."
            .VerticalGridStyle = xtpGridSolid
        End With
        Set Column = .Add(mRCol.������Ŀ, "������Ŀ", 85, True)
        Set Column = .Add(mRCol.���, "���", 85, True)
        Set Column = .Add(mRCol.��λ, "��λ", 65, True)
        Set Column = .Add(mRCol.��־, "��־", 65, True)
        Set Column = .Add(mRCol.�ο�, "�ο�", 85, True)
    End With
End Sub

Private Sub Form_Resize()
'    With Me.rptList
'        .Left = 0
'        .Top = 0
'        .Width = Me.ScaleWidth
'        .Height = Me.ScaleHeight
'    End With
End Sub
Public Function zlRefresh(ByVal lngSampleID As Long, strName As String, strSex As String, strAge As String, ItemName As String, _
                        lngPatientID As String, strMachineName As String, SampleTime As String, strVerifyName As String) As Boolean
    Dim Record As ReportRecord
    Dim intLoop As Integer
    Dim intRowIndex As Integer
    
    zlRefresh = True
    
    intRowIndex = CheckName
    
    If intRowIndex > 0 And strName <> "" Then
        If MsgBox("�Ѵ���һ���ϲ��걾,�Ƿ񸲸�?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            
            zlRefresh = False
            Exit Function
        End If
        Me.rptList.Records.RemoveAt intRowIndex - 1
    End If
    
    For intLoop = 1 To Me.rptList.Records.Count
        If Me.rptList.Records(intLoop - 1).Item(mCol.�걾ID).Value = lngSampleID Then
            MsgBox "���Ѽ����˸�ͬ���ı걾!", vbQuestion, gstrSysName
            zlRefresh = False
            Exit Function
        End If
    Next
    
    Set Record = Me.rptList.Records.Add
    
    For intLoop = 0 To Me.rptList.Columns.Count
        Record.AddItem ""
    Next
    
    Record.Item(mCol.�걾ID).Value = lngSampleID
    Record.Item(mCol.����).Value = strName
    Record.Item(mCol.�Ա�).Value = strSex
    Record.Item(mCol.����).Value = strAge
    Record.Item(mCol.������Ŀ).Value = ItemName
    Record.Item(mCol.��ʶ��).Value = lngPatientID
    Record.Item(mCol.����).Value = strMachineName
    Record.Item(mCol.�걾ʱ��).Value = SampleTime
    Record.Item(mCol.������).Value = strVerifyName
    
    If strName <> "" Then
        '����һ��ʱĬ��Ϊ�ϲ��걾
        Record.Item(mCol.�ϲ�״̬).Value = "�ϲ��걾"
        For intLoop = 1 To Me.rptList.Columns.Count
            Record.Item(intLoop).ForeColor = vbBlue
        Next
    Else
        Record.Item(mCol.�ϲ�״̬).Value = "���ϲ��걾"
        For intLoop = 1 To Me.rptList.Columns.Count
            Record.Item(intLoop).ForeColor = vbRed
        Next
    End If
    
    Me.rptList.Populate
    
    RefreshUnion
      
End Function

Private Sub Form_Unload(Cancel As Integer)
    Me.cbrthis.DeleteAll
    Me.dkpMain.DestroyAll
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, Y As Long)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    
    If Button <> vbRightButton Then Exit Sub
    
    If Me.rptList.Records.Count = 0 Then Exit Sub

    Set cbrPopupBar = Me.cbrthis.Add("�����˵�", xtpBarPopup)
'    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Insert, "���úϲ��걾")
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_ThingDel, "������кϲ��걾")
    
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Me.rptList.Records.RemoveAt (Row.Index)
    Me.rptList.Populate
    Call RefreshUnion
End Sub

Public Function ZlSave() As Long
    Dim intLoop As Integer
    Dim lngSourceID As Long
    Dim strUnionID As String
    Dim varID() As String
    
    On Error GoTo errH
    
    '�Ƿ���������걾
    If Me.rptList.Records.Count < 2 Then
        MsgBox "�걾�����������ܺϲ���ѡ��걾���ٵ�ϲ�!", vbQuestion, gstrSysName
        Exit Function
    End If
    
    '�Ƿ��������걾
    If CheckName = 0 Then
        MsgBox "��ѡ��һ�������걾��,�ٺϲ�!", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�õ��ϲ�ID
    For intLoop = 1 To Me.rptList.Records.Count
        If rptList.Records(intLoop - 1).Item(mCol.�ϲ�״̬).Value = "�ϲ��걾" Then
            lngSourceID = Me.rptList.Records(intLoop - 1).Item(mCol.�걾ID).Value
        Else
            strUnionID = strUnionID & ";" & Me.rptList.Records(intLoop - 1).Item(mCol.�걾ID).Value
        End If
    Next
    strUnionID = Mid(strUnionID, 2)
    
    
                        
    If lngSourceID > 0 And strUnionID <> "" Then
        gcnOracle.BeginTrans
        
        varID = Split(strUnionID, ";")
        
        For intLoop = 0 To UBound(varID)
            gstrSql = "Zl_����걾��¼_Union(" & lngSourceID & "," & varID(intLoop) & ")"
            zlDatabase.ExecuteProcedure gstrSql, gstrSysName
        Next
        
        gcnOracle.CommitTrans
    End If
    
    Me.rptList.Records.DeleteAll
    Me.rptList.Populate
    Call RefreshUnion
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Function
Private Function CheckName() As Integer
    '�����б����Ƿ����в��������ļ�¼����
    Dim intLoop As Integer
    
    For intLoop = 1 To Me.rptList.Records.Count
        If Me.rptList.Records(intLoop - 1).Item(mCol.����).Value <> "" Then
            CheckName = intLoop
            Exit For
        End If
    Next


End Function

Private Sub RefreshUnion()
    '����       ˢ�ºϲ��ͱ��ϲ����
    Dim rsTmp As New adodb.Recordset
    Dim lngKey As Long      '�ϲ��걾ID
    Dim lngUnionKey As Long '���ϲ��걾ID
    Dim intLoop As Integer
    Dim Record As ReportRecord

    
    Me.RptComPare1.Records.DeleteAll
    Me.RptComPare2.Records.DeleteAll
    Me.RptComPare1.Populate
    Me.RptComPare2.Populate
    If Me.rptList.Rows.Count = 0 Then Exit Sub
    
    
    With Me.rptList
        For intLoop = 0 To .Rows.Count - 1
            If .Rows(intLoop).Record(mCol.�ϲ�״̬).Value = "�ϲ��걾" Then
                lngKey = Val(.Rows(intLoop).Record(mCol.�걾ID).Value)
            End If
            If .Rows(intLoop).Record(mCol.�ϲ�״̬).Value = "���ϲ��걾" Then
                lngUnionKey = Val(.Rows(intLoop).Record(mCol.�걾ID).Value)
                If .Rows(intLoop).Selected = True Then
                    Exit For
                End If
            End If
            
        Next
    End With
    
    
    '�ϲ�
    If lngKey > 0 Then
        gstrSql = "Select C.������ || Decode(C.Ӣ����, Null, '', '(' || C.Ӣ���� || ')') As ������Ŀ, B.������, C.��λ," & vbNewLine & _
                "       Trim(Replace(Replace(' ' || Zlgetreference(C.ID, A.�걾����, Decode(A.�Ա�, '��', 1, 'Ů', 2, 0), A.��������, A.����id, A.����,a.�������id)," & vbNewLine & _
                "                             ' .', '0.'), '��.', '��0.')) As �ο�, " & vbNewLine & _
                " DECODE(B.�����־,3,'��',2,'��',1,'',4,'�쳣',5,'����',6,'����','') AS ��־,B.�����־ " & vbNewLine & _
                "From ����걾��¼ A, ������ͨ��� B, ����������Ŀ C, ������ĿĿ¼ D, ������Ŀ E" & vbNewLine & _
                "Where A.ID = B.����걾id And B.������Ŀid = C.ID And B.������Ŀid = D.ID(+) And B.������Ŀid = E.������Ŀid And A.ID = [1]" & vbNewLine & _
                "Order By Decode(E.�������, Null, Nvl(D.����, 9999999999), E.�������), B.������� "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKey)
            
        Do Until rsTmp.EOF
            Set Record = Me.RptComPare1.Records.Add
            For intLoop = 0 To Me.RptComPare1.Columns.Count
                Record.AddItem ""
            Next
            Record(mRCol.������Ŀ).Value = Nvl(rsTmp("������Ŀ"))
            Record(mRCol.���).Value = Nvl(rsTmp("������"))
            Record(mRCol.��־).Value = Nvl(rsTmp("��־"))
            Record(mRCol.��λ).Value = Nvl(rsTmp("��λ"))
            Record(mRCol.�ο�).Value = Nvl(rsTmp("�ο�"))
            Call ApplyResultColor(Record, Val(Nvl(rsTmp("�����־"))))
            rsTmp.MoveNext
        Loop
    End If
    
    If lngKey > 0 And lngUnionKey > 0 Then
        gstrSql = "Select C.������ || Decode(C.Ӣ����, Null, '', '(' || C.Ӣ���� || ')') As ������Ŀ, B.������, C.��λ," & vbNewLine & _
                    "       Trim(Replace(Replace(' ' || Zlgetreference(C.ID, F.�걾����, Decode(F.�Ա�, '��', 1, 'Ů', 2, 0), F.��������, F.����id, F.����,f.�������id)," & vbNewLine & _
                    "                             ' .', '0.'), '��.', '��0.')) As �ο�, " & vbNewLine & _
                    " DECODE(B.�����־,3,'��',2,'��',1,'',4,'�쳣',5,'����',6,'����','') AS ��־,B.�����־ " & vbNewLine & _
                    "From ����걾��¼ A, ������ͨ��� B, ����������Ŀ C, ������ĿĿ¼ D, ������Ŀ E, ����걾��¼ F" & vbNewLine & _
                    "Where A.ID = B.����걾id And B.������Ŀid = C.ID And B.������Ŀid = D.ID(+) And B.������Ŀid = E.������Ŀid And F.ID = [2] And" & vbNewLine & _
                    "      A.ID = [1]" & vbNewLine & _
                    "Order By Decode(E.�������, Null, Nvl(D.����, 9999999999), E.�������), B.�������"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngUnionKey, lngKey)
            
        Do Until rsTmp.EOF
            Set Record = Me.RptComPare2.Records.Add
            For intLoop = 0 To Me.RptComPare2.Columns.Count
                Record.AddItem ""
            Next
            Record(mRCol.������Ŀ).Value = Nvl(rsTmp("������Ŀ"))
            Record(mRCol.���).Value = Nvl(rsTmp("������"))
            Record(mRCol.��־).Value = Nvl(rsTmp("��־"))
            Record(mRCol.��λ).Value = Nvl(rsTmp("��λ"))
            Record(mRCol.�ο�).Value = Nvl(rsTmp("�ο�"))
            Call ApplyResultColor(Record, Val(Nvl(rsTmp("�����־"))))
            rsTmp.MoveNext
        Loop

    End If
    
    Me.RptComPare1.Populate
    Me.RptComPare2.Populate
End Sub


Private Sub rptList_SelectionChanged()
    Call RefreshUnion
End Sub
Private Sub ApplyResultColor(Record As ReportRecord, bytMode As Byte)
    '-----------------------------------------------------------------------------------------
    '����:
    '-----------------------------------------------------------------------------------------
    Dim lngColor As Long, lngForeColor As Long
    
    Select Case bytMode
        Case 0, 1
            lngColor = vbWhite
            lngForeColor = COLOR.Ĭ��ǰ��ɫ
        Case 5, 6 '�쳣�͡���
            lngColor = COLOR.��������ɫ
            lngForeColor = vbWhite
        Case 2
            lngColor = COLOR.�ͱ걳��ɫ
            lngForeColor = COLOR.����ǰ��ɫ
        Case Else
            lngColor = COLOR.���걳��ɫ
            lngForeColor = COLOR.����ǰ��ɫ
    End Select
    
    Record.Item(mRCol.���).BackColor = lngColor
    Record.Item(mRCol.���).ForeColor = lngForeColor
End Sub
