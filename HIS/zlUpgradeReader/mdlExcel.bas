Attribute VB_Name = "mdlExcel"
Option Explicit

Public gcnExcel As New ADODB.Connection '��������

Public Enum Excel_Col
    'excel��ʽ�ļ����ֶ�˳��
    Ӱ������ = 0
    �����汾
    �Ƿ���HTML�ĵ�
    ������
    �Ǽ�ģ��
    �Ǽ��û�
    �û�����
    �޸�˵��
    �������
    ��������˵��
    Ӱ��ģ��
    �Ƿ���Ҫ��ѵ
    ��ע
    �������
    �Ķ���¼
    ���û�Ӱ������
    ������ѵ���
End Enum

Public Enum mCol
    �Ķ� = 0: ����: �汾: ����: ���: ģ��: Ӱ��ģ��: Ӱ������: ��������: �û�: ����: ˵��: ��������: ��ע: ��ѵ: Ӱ������: ����: �޸�
End Enum
Public fntUnderLine  As StdFont '�����ӵ�����


Public Function OpenExcelFile(ByVal strFilename As String) As String
    '���ܣ���Excel��ʽ�ļ�
    '��Σ�strFileName
    '���Σ�Sheet�б���|�ָ�
    
    Dim BiaoMing As Variant
    Dim TableName As String
    Dim strSheet As String
    On Error GoTo errHandle
    OpenExcelFile = ""

    If gcnExcel.State = 1 Then     '��������ӹ�����رգ���ʼ���´�����
        gcnExcel.Close
    End If
    
    gcnExcel.ConnectionString = "Provider=microsoft.jet.oledb.4.0;data source=" & strFilename & ";" & _
                              "Extended Properties=Excel 8.0;" & _
                              "Persist Security Info=False"
    gcnExcel.Open
    Set BiaoMing = gcnExcel.OpenSchema(adSchemaTables)     '�������ݿ��¼��
    
    TableName = "": strSheet = ""
    Do Until BiaoMing.EOF
        If BiaoMing("table_name") <> TableName Then   '�г����б�
            TableName = BiaoMing("table_name")
            If Right(TableName, 1) = "$" Then
                strSheet = strSheet & "|" & TableName
            End If
        End If
        BiaoMing.MoveNext
    Loop
    
    Set BiaoMing = Nothing
    If strSheet <> "" Then
        OpenExcelFile = Mid(strSheet, 2)
    End If
    Exit Function
errHandle:
    OpenExcelFile = ""
    MsgBox Err.Number & " " & Err.Description, vbQuestion, "�����Ķ���"
    
End Function

Public Function OpenExcelSheet(ByVal strSheetName As String) As ADODB.Recordset
    '��һ��Sheet
    '���: Sheet��
    '����: ADO��¼��
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSheet As String
    On Error GoTo errHandle
    
    If strSheetName = "" Then Exit Function
    
    strSheet = strSheetName
    If Right(strSheet, 1) <> "$" Then
        strSheet = strSheet & "$"
    End If
    
    rsTmp.Open strSheetName, gcnExcel, adOpenDynamic, adLockPessimistic, adCmdTableDirect
    If Not rsTmp.EOF Then
        Set OpenExcelSheet = rsTmp
    End If

    Exit Function
errHandle:
    If Err.Number = -2147217865 Then Exit Function
    MsgBox Err.Number & " " & Err.Description, vbQuestion, "�����Ķ���"
End Function


Public Sub initRptList(ByRef objRpt As ReportControl, ByRef objImg As ImageList, ByVal txtFont As StdFont, ByVal blnEdit As Boolean)
    '��ʼ��report�ؼ�
    
    Dim rptCol As ReportColumn

    Dim TextFont As StdFont
    '��ʼ���б�
    
    With objRpt
        .SetImageList objImg
        
        '.AutoColumnSizing = (Screen.Width / Screen.TwipsPerPixelX > 800)   '������������֮ǰ���ã�������Ч
        '�Ѷ� = 0: ����: �汾: ����: ���: ģ��: Ӱ��ģ��: ��������: �û�:����: ˵��: ��������: ��ע:Ӱ������: ��ѵ:  ����
        Set rptCol = .Columns.Add(mCol.�Ķ�, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        rptCol.Icon = ICON_Mail
        
        Set rptCol = .Columns.Add(mCol.����, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        rptCol.Icon = ICON_Importance
        
        Set rptCol = .Columns.Add(mCol.�汾, "�汾", 60, True): rptCol.Editable = False: rptCol.Groupable = True
        
        
        Set rptCol = .Columns.Add(mCol.����, "��������", 50, True): rptCol.Editable = False: rptCol.Groupable = True
        
        
        Set rptCol = .Columns.Add(mCol.���, "���", 60, True): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        
        
        Set rptCol = .Columns.Add(mCol.ģ��, "ģ��", 120, True): rptCol.Editable = False: rptCol.Groupable = True
        
        
        Set rptCol = .Columns.Add(mCol.Ӱ��ģ��, "Ӱ��ģ��", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        
        Set rptCol = .Columns.Add(mCol.Ӱ������, "Ӱ������", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        
        Set rptCol = .Columns.Add(mCol.��������, "��������", 10, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        
        Set rptCol = .Columns.Add(mCol.�û�, "�Ǽ��û�", 80, True): rptCol.Editable = False: rptCol.Groupable = True
        
        
        Set rptCol = .Columns.Add(mCol.����, "�û�����", 10, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        
        
        Set rptCol = .Columns.Add(mCol.˵��, "�޸�˵��", 10, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        
        
        Set rptCol = .Columns.Add(mCol.��������, "��������", 80, True): rptCol.Editable = False: rptCol.Groupable = False
        
        Set rptCol = .Columns.Add(mCol.��ע, "��ע", 80, True): rptCol.Editable = False: rptCol.Groupable = False
        
        Set rptCol = .Columns.Add(mCol.��ѵ, "��ѵ���", 70, False): rptCol.Editable = False: rptCol.Groupable = False
        'rptCol.Icon = ICON_FlagTrain
        
        Set rptCol = .Columns.Add(mCol.Ӱ������, "Ӱ������", 70, True): rptCol.Editable = blnEdit: rptCol.Groupable = False
        With rptCol.EditOptions
            .Constraints.Add "", δ��д
            .Constraints.Add "��������", ��������
            .Constraints.Add "��������", ��������
            .Constraints.Add "��Ӱ��", ��Ӱ��
            .ConstraintEdit = True
            .AddComboButton
        End With
        
        Set rptCol = .Columns.Add(mCol.����, "����", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.�޸�, "�޸�", 0, False): rptCol.Editable = True: rptCol.Groupable = False: rptCol.Visible = False
        
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .ShowGroupBox = True
        With .PaintManager
            
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 2
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            '.VerticalGridStyle = xtpGridSolid
            '��������
            Set TextFont = txtFont
            'TextFont.Size = 12
            Set .TextFont = TextFont
            Set .CaptionFont = TextFont
            
            '����������
            Set fntUnderLine = .TextFont
            fntUnderLine.Underline = True
                        
        End With
        .PreviewMode = False
        .AllowEdit = True
        .EditOnClick = True
        .FocusSubItems = True
    
        '�������
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(mCol.����)
        .GroupsOrder(0).SortAscending = True
        .Columns.Find(mCol.����).Visible = False
        .Populate
    End With
End Sub
