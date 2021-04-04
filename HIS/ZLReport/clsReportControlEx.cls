VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReportControlEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------
'��;��ReportControl�ؼ�����չ��װ
'˵����
'   Recordset->Bunding->ImageList->Init->��NodeName��NodeNameParent��NodeTree��->Repaint
'���ߣ�YuZhiYong
'���䣺yzyong@zlsoft.cn��154174331@qq.com
'-----------------------------------------

Public Enum ViewMode_Enum
    vmListView = 0      'ListViewģʽ
    vmTreeView = 1      'TreeViewģʽ
End Enum

Public Enum AppTemplate_Enum
    atTree = 0
    atTreeEx = 1
    atList = 2
    atGroup = 3
End Enum

Public Event DragDropEx(Source As Control, X As Single, Y As Single)

Private Const GSTR_MSG As String = "clsReportControlEx��"

Private WithEvents mobjBinding As XtremeReportControl.ReportControl         '�󶨱�׼�ؼ�
Attribute mobjBinding.VB_VarHelpID = -1
Private mfrmOwner As Form
Private mblnDragBegin As Boolean

Private menuTemplate As AppTemplate_Enum                                    'ҵ��Ӧ��ģ��
Private menuViewMode As ViewMode_Enum                                       'ReportControl��ʾģʽ
Private mobjList As ImageList
Private mstrNodeName As String                                              '����Ӧ��Recordset�ֶ���
Private mstrNodeNameParent As String                                        '������Ӧ��Recordset�ֶ���
Private mstrNodeTree As String                                              '��������Ӧ��Recordset�ֶ���
Private mstrColumnHide As String                                            '��ʾ���ֶ���
Private mblnAllowDrag As Boolean                                            '�����Ϸ�
Private mlngIconFolder As Long
Private mlngIconNode As Long

Public Property Get IconNode() As Long
    IconNode = mlngIconNode
End Property

Public Property Let IconNode(ByVal LonValue As Long)
    mlngIconNode = LonValue
End Property

Public Property Get IconFolder() As Long
    IconFolder = mlngIconFolder
End Property

Public Property Let IconFolder(ByVal LonValue As Long)
    mlngIconFolder = LonValue
End Property

Public Property Get AllowDrag() As Boolean
    AllowDrag = mblnAllowDrag
End Property

Public Property Let AllowDrag(ByVal blnValue As Boolean)
    mblnAllowDrag = blnValue
    If mblnAllowDrag Then
        If mobjBinding Is Nothing Then
            MsgBox "����ʹ��Binding���԰󶨿ؼ���", vbInformation, GSTR_MSG
        Else
            mobjBinding.MultipleSelection = False      '�����Ϸţ��Ͳ�֧�ֶ�ѡ
        End If
    End If
End Property

Public Property Get ColumnHide() As String
'��ʽ���ֶ���1[|�ֶ���2[...]]
    ColumnHide = mstrColumnHide
End Property

Public Property Let ColumnHide(ByVal strValue As String)
    mstrColumnHide = strValue
End Property

Public Property Get NodeTree() As String
    NodeTree = mstrNodeTree
End Property

Public Property Let NodeTree(ByVal strValue As String)
    mstrNodeTree = strValue
End Property

Public Property Get NodeNameParent() As String
    NodeNameParent = mstrNodeNameParent
End Property

Public Property Let NodeNameParent(ByVal strValue As String)
    mstrNodeNameParent = strValue
End Property

Public Property Get NodeName() As String
    NodeName = mstrNodeName
End Property

Public Property Let NodeName(ByVal strValue As String)
    mstrNodeName = strValue
End Property

Public Property Get ImageList() As ImageList
    ImageList = mobjList
End Property

Public Property Let ImageList(ByVal ObjValue As ImageList)
    Set mobjList = ObjValue
    If Not mobjList Is Nothing Then
        If menuTemplate = atTreeEx Then
            If mobjList.ListImages("Folder") Is Nothing Or mobjList.ListImages("Node") Is Nothing Then
                MsgBox "ImageList���������ڡ�Folder���͡�Node������ͼ����Ŀ��", vbCritical, GSTR_MSG
                Exit Property
            End If
        Else
            If mobjList.ListImages("Folder") Is Nothing Then
                MsgBox "ImageList���������ڡ�Folder��ͼ����Ŀ��", vbCritical, GSTR_MSG
                Exit Property
            End If
        End If
    End If
End Property

Public Property Get Binding() As XtremeReportControl.ReportControl
    Set Binding = mobjBinding
End Property

Public Property Let Binding(ByVal ObjValue As XtremeReportControl.ReportControl)
    Set mobjBinding = ObjValue
End Property

Public Property Get ViewMode() As ViewMode_Enum
    ViewMode = menuViewMode
End Property

Public Property Let ViewMode(ByVal BytValue As ViewMode_Enum)
    menuViewMode = BytValue
End Property

Public Sub AppTemplate(ByVal enuTemplate As AppTemplate_Enum _
    , ByVal objBinding As XtremeReportControl.ReportControl _
    , Optional ByVal objImageList As ImageList = Nothing _
    , Optional ByVal strColumnHide As String = "" _
    , Optional ByVal strTreeInfo As String = "" _
    , Optional ByVal lngIconFolder As Long = 0 _
    , Optional ByVal lngIconNode As Long = 0 _
)
'���ܣ�ҵ��Ӧ��ģ��
'������
'  TreeInfo�����ͽṹ��Ϣ����ʽ������ֶ���|������ֶ���|�����������

    Dim arrInfo As Variant
    Dim i As Integer

    menuTemplate = enuTemplate
    Binding = objBinding
    ImageList = objImageList
    ColumnHide = strColumnHide

    '���ͽṹ
    Select Case menuTemplate
    Case atTree, atTreeEx
        ViewMode = vmTreeView
        IconFolder = lngIconFolder
        IconNode = lngIconNode

        arrInfo = Split(strTreeInfo, "|")

        If UBound(arrInfo) = -1 Then
            MsgBox "�봫�롰strTreeInfo��������", vbCritical, GSTR_MSG
            Exit Sub
        End If
        If UBound(arrInfo) <> 2 Then
            MsgBox "�봫����ȷ�ġ�strTreeInfo��������", vbCritical, GSTR_MSG
            Exit Sub
        End If

        For i = LBound(arrInfo) To UBound(arrInfo)
            Select Case i
            Case Val("0-����ֶ���")
                NodeName = arrInfo(i)
            Case Val("1-������ֶ���")
                NodeNameParent = arrInfo(i)
            Case Val("2-�����������")
                NodeTree = arrInfo(i)
            End Select
        Next
    Case atGroup
        ViewMode = vmListView
    Case menuTemplate
        ViewMode = vmListView
    End Select
        
End Sub

Public Sub Init(ByVal frmOwner As Form)
'���ܣ���ʼ���ؼ��Ļ�������
'������
'  frmOwner�������������
    
    Set mfrmOwner = frmOwner
    
    With mobjBinding
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GroupForeColor = &HC00000
            If menuTemplate = atTree Or menuTemplate = atTreeEx Then
                .GridLineColor = .BackColor
            Else
                .GridLineColor = RGB(225, 225, 225)
            End If
            .VerticalGridStyle = xtpGridSolid
            .NoItemsText = "û�п���ʾ����Ŀ..."
        End With
        
        If ViewMode = vmTreeView Then
            .AllowColumnSort = False
        End If
        
        .PreviewMode = False
        .AutoColumnSizing = True
        .AllowColumnRemove = False
        .AllowColumnResize = True
        '.BorderStyle = xtpBorderFlat
    End With
End Sub

Public Sub Repaint(ByVal rsData As ADODB.Recordset, Optional ByVal blnColumn As Boolean = False)
'���ܣ���������
'  rsData����
'  blnColumn��True-������ͷ���ݣ�False-���Ʊ�������

    Dim objRec As XtremeReportControl.ReportRecord
    Dim objCol As XtremeReportControl.ReportColumn
    Dim i As Long, lngIdx As Long, lngOldRow As Long
    
    '���
    If ViewMode = vmTreeView Then
        If NodeName = "" Then
            MsgBox "�����á�NodeName�����ԣ�", vbCritical, GSTR_MSG
            Exit Sub
        End If
        If NodeNameParent = "" Then
            MsgBox "�����á�NodeNameParent�����ԣ�", vbCritical, GSTR_MSG
            Exit Sub
        End If
        If NodeTree = "" Then
            MsgBox "�����á�NodeTree�����ԣ�", vbCritical, GSTR_MSG
            Exit Sub
        End If
    End If
    
    '����
    With mobjBinding
        If blnColumn Then
            '��ͷ����
            For i = 0 To rsData.Fields.count - 1
                Set objCol = .Columns.Add(i, rsData(i).name, 0, True)
                objCol.Editable = False
                objCol.Groupable = False
                objCol.Sortable = False
                objCol.Visible = Not ("|" & ColumnHide & "|" Like "*|" & objCol.Caption & "|*")
                objCol.AllowDrag = True
            Next
            .Columns(GetColIndex(NodeTree)).TreeColumn = True
        Else
            lngOldRow = .FocusedRow.Index
            .Records.DeleteAll
        End If
            
        '���ݼ���
        rsData.Filter = ""
        Do While rsData.EOF = False
            Call FillRecord(rsData.Fields)
            rsData.MoveNext
        Loop
        
        .Populate
        
        '����ͼ��
        If menuTemplate = atTree Or menuTemplate = atTreeEx Then
            lngIdx = GetColIndex(NodeTree)
            For i = 0 To .Rows.count - 1
                If menuTemplate = atTreeEx Then
                    If .Rows(i).Record.Childs.count > 0 Then
                        .Rows(i).Record(lngIdx).Icon = IconFolder
                    Else
                        .Rows(i).Record(lngIdx).Icon = IconNode
                    End If
                Else
                    .Rows(i).Record(lngIdx).Icon = IconFolder
                End If
            Next
        End If
    
        If lngOldRow > 0 Then
            If lngOldRow > .Rows.count - 1 Then
                .Rows(.Rows.count - 1).Selected = True
            Else
                .Rows(lngOldRow).Selected = True
            End If
        End If
        
    End With
End Sub

Private Function FindRecord(ByVal strTag As String _
    , ByVal objRECs As XtremeReportControl.ReportRecords) As XtremeReportControl.ReportRecord
'���ܣ��ݹ����ָ��Tagֵ�Ķ��󣬼�Ψһ��¼�����ֵ
'������
'  strTag��
'  objRECs��

    Dim l As Long
    
    For l = objRECs.count - 1 To 0 Step -1
        If strTag = objRECs(l).Tag Then
            Set FindRecord = objRECs(l)
            Exit Function
        Else
            If objRECs(l).Childs.count > 0 Then
                Set FindRecord = FindRecord(strTag, objRECs(l).Childs)
                Exit Function
            Else
                If strTag = objRECs(l).Tag Then
                End If
            End If
        End If
    Next
End Function

Private Function FillRecord(ByVal objFields As ADODB.Fields) As XtremeReportControl.ReportRecord
    Dim objRec As XtremeReportControl.ReportRecord, objParent As XtremeReportControl.ReportRecord
    Dim objCell As XtremeReportControl.ReportRecordItem
    Dim i As Integer
    Dim strParentTag As String
    
    If objFields Is Nothing Then Exit Function
    
    If IsNull(objFields(NodeNameParent).Value) Then
        strParentTag = "0"
    Else
        strParentTag = "" & objFields(NodeNameParent).Value
    End If
    
    '���Ҹ�����
    Set objParent = FindRecord(strParentTag, mobjBinding.Records)
    
    '��������
    If objParent Is Nothing Then
        Set objRec = mobjBinding.Records.Add
    Else
        Set objRec = objParent.Childs.Add
    End If
    objRec.Tag = CStr(objFields(NodeName).Value)
    
    For i = 0 To objFields.count - 1
        Set objCell = objRec.AddItem(objFields(i).Value)
    Next
    objRec.Expanded = True
    
    Set FillRecord = objRec
End Function

Public Function GetColIndex(ByVal strColName As String) As Integer
    Dim i As Integer
    
    For i = 0 To mobjBinding.Columns.count - 1
        If UCase(strColName) = UCase(mobjBinding.Columns(i).Caption) Then
            GetColIndex = i
            Exit Function
        End If
    Next
End Function

Private Sub mobjBinding_DragDrop(Source As Control, X As Single, Y As Single)
    Dim objInfo As XtremeReportControl.ReportHitTestInfo
    Dim objRec As XtremeReportControl.ReportRecord
    Dim objCell As XtremeReportControl.ReportRecordItem
    Dim i As Integer
    Dim lngRowIns, lngRowDel As Long
    Dim strTag As String
    
    If Source.name = mobjBinding.name Then
        '�����Ϸ�
        Set objInfo = mobjBinding.HitTest(mfrmOwner.ScaleX(X, vbTwips, vbPixels), mfrmOwner.ScaleY(Y, vbTwips, vbPixels))
        If Not objInfo.Row Is Nothing Then
            If objInfo.Row.Index = mobjBinding.FocusedRow.Index Then Exit Sub
        
            lngRowIns = objInfo.Row.Index
            lngRowDel = mobjBinding.FocusedRow.Index
            
            '��λ�ò���
            If objInfo.Row.ParentRow Is Nothing Then
                Set objRec = objInfo.Row.Record.Childs.Insert(objInfo.Row.Record.Index)
            Else
                Set objRec = objInfo.Row.ParentRow.Record.Childs.Insert(objInfo.Row.Record.Index)
            End If
            For i = 0 To mobjBinding.Columns.count - 1
                Set objCell = objRec.AddItem(mobjBinding.FocusedRow.Record(i).Value)
                objCell.Icon = mobjBinding.FocusedRow.Record(i).Icon
            Next
            strTag = mobjBinding.FocusedRow.Record.Tag
                            
            'ԭλ��ɾ��
            Call DelRec(mobjBinding.Records, mobjBinding.FocusedRow.Record.Tag)
            
            '����
            objRec.Tag = strTag
            
            'ˢ��
            mobjBinding.Populate
            
            '�������󽹵�
            If lngRowIns < lngRowDel Then
                mobjBinding.FocusedRow = mobjBinding.Rows(lngRowIns)
            ElseIf lngRowIns > lngRowDel Then
                mobjBinding.FocusedRow = mobjBinding.Rows(lngRowIns - 1)
            End If
        End If
    Else
        '�����ؼ�����������¼�������ҵ����أ�
        'RaiseEvent DragDropEx(Source, X, Y)
    End If
End Sub

Private Sub mobjBinding_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Dim objInfo As ReportHitTestInfo
    
    Set objInfo = mobjBinding.HitTest(mfrmOwner.ScaleX(X, vbTwips, vbPixels), mfrmOwner.ScaleY(Y, vbTwips, vbPixels))
    If Not objInfo.Row Is Nothing Then
        If objInfo.Row.Selected = False Then objInfo.Row.Selected = True
    End If
End Sub

Private Sub mobjBinding_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    mblnDragBegin = False
End Sub

Private Sub mobjBinding_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button = vbLeftButton Then
        If mblnAllowDrag And Not mobjBinding.FocusedRow Is Nothing Then
            '����reportcontrol.drag��mousemove�¼����ᱻ����
            If mobjBinding.FocusedRow.Record.Childs.count <= 0 Then
                If mblnDragBegin = False Then
                    '�����Ϸ�
                    Call mobjBinding.Drag
                    'Set mobjBinding.DragIcon = imgCustom.ListImages("Node").Picture 'ͼ�������������
                    mblnDragBegin = True
                End If
            End If
        End If
    End If
End Sub

Private Sub mobjBinding_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    mblnDragBegin = False
End Sub

Public Sub DelRec(ByVal objRECs As XtremeReportControl.ReportRecords, Optional strTag As String)
    Dim l As Long
    
    For l = objRECs.count - 1 To 0 Step -1
        If objRECs(l).Childs.count <= 0 Then
            If strTag = objRECs(l).Tag Then
                objRECs.RemoveAt objRECs(l).Index
                Exit Sub
            End If
        Else
            Call DelRec(objRECs(l).Childs, strTag)
            If objRECs(l).Childs.count <= 0 Then
                objRECs.RemoveAt objRECs(l).Index
            End If
        End If
    Next
End Sub