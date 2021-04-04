VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVSFlexGridEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------
'��;����׼VSFlexGrid����չ��װ
'˵����
'1.����ִ��˳��
'      Recordset->Binding->Init->Head->ColsReadonly->Editable->Repaint
'���ߣ�YuZhiYong
'���䣺yzyong@zlsoft.cn
'-----------------------------------------

Private Const GSTR_MSG As String = "clsVSFlexGridEx��"

'�󶨵�VSFlexGrid�ؼ�
Private WithEvents mvsfBinding As VSFlexGrid
Attribute mvsfBinding.VB_VarHelpID = -1

'��������¼�
Public Event EventFillData(ByVal vsfVar As VSFlexGrid, ByVal Row As Long, ByVal Col As Long)

Public Enum enuEditable
    EM_View = 0
    EM_Modify = 1
    EM_Verify = 2
    EM_Display = 3
End Enum

Public Enum enuRepaintType
    RT_Columns = 0
    RT_Rows = 1
    RT_ColsAndRows = 2
End Enum

Private Type typSort
    Col As Integer
    type As Integer
End Type
Private mtypSort As typSort

'�󶨵ļ�¼��������չ�֣�
Private mrsRecordset As ADODB.Recordset

'��ͷ
Private mstrHead As String
'ֻ�����ַ�����������1[|��2[|...]]
Private mstrColsReadonly As String
'�ؼ�ģʽ
Private mbytEditable As enuEditable
'Append����Ҫ������һ���е�ֵ
Private mstrAppendFields As String
'�������ַ�����������1[|��2[|...]]
Private mstrColsHide As String
'������ѡ
Private mblnAllowMultSelection As Boolean
'�и�
Private mintHeightRow As Integer
'��ͷ��
Private mintHeightColumn As Integer

Public Property Get HeightColumn() As Integer
    HeightColumn = mintHeightColumn
End Property

Public Property Let HeightColumn(ByVal intValue As Integer)
    mintHeightColumn = intValue
End Property

Public Property Get HeightRow() As Integer
    HeightRow = mintHeightRow
End Property

Public Property Let HeightRow(ByVal intValue As Integer)
    mintHeightRow = intValue
End Property

Public Property Get AllowMultSelection() As Boolean
    AllowMultSelection = mblnAllowMultSelection
End Property

Public Property Let AllowMultSelection(ByVal bValue As Boolean)
    mblnAllowMultSelection = bValue
End Property

Public Property Get ColsHide() As String
    ColsHide = mstrColsHide
End Property

Public Property Let ColsHide(ByVal strValue As String)
    mstrColsHide = strValue
End Property

Public Property Get AppendFields() As String
    AppendFields = mstrAppendFields
End Property

Public Property Let AppendFields(ByVal strValue As String)
    mstrAppendFields = strValue
End Property

Public Property Get Editable() As enuEditable
    Editable = mbytEditable
End Property

Public Property Let Editable(ByVal BytValue As enuEditable)
    mbytEditable = BytValue
    If BytValue = EM_Modify Or BytValue = EM_Verify Then
        mvsfBinding.Editable = flexEDKbdMouse
    Else
        mvsfBinding.Editable = flexEDNone
    End If
End Property

Public Property Get ColsReadonly() As String
    ColsReadonly = mstrColsReadonly
End Property

Public Property Let ColsReadonly(ByVal strValue As String)
    mstrColsReadonly = strValue
End Property

Public Property Get Head() As String
    Head = mstrHead
End Property

Public Property Let Head(ByVal strValue As String)
    mstrHead = strValue
    '����������
    Call SetVSFHead(mvsfBinding, mstrHead)
End Property

Public Property Get Recordset() As ADODB.Recordset
    Set Recordset = mrsRecordset
End Property

Public Property Let Recordset(ByVal RecValue As ADODB.Recordset)
    Set mrsRecordset = RecValue
End Property

Public Property Get Binding() As VSFlexGrid
    Set Binding = mvsfBinding
End Property

Public Property Let Binding(ByVal VSFValue As VSFlexGrid)
    Set mvsfBinding = VSFValue
End Property

Public Sub AppTemplate(ByVal enuMode As enuEditable _
    , ByVal objBinding As VSFlexGrid _
    , ByVal strColumns As String _
    , ByVal strColumnsReadonly As String _
    , Optional ByVal blnMultSelection As Boolean = False _
)
    Binding = objBinding
    Editable = enuMode
    Head = strColumns
    ColsReadonly = strColumnsReadonly
    mblnAllowMultSelection = blnMultSelection
    
    Call Init
End Sub

Public Sub Repaint(Optional ByVal typVar As enuRepaintType)
'���ܣ��ػ�VSFlexGrid�ؼ�
'������
'  typVar���ػ�Grid

    If typVar = RT_Columns Or typVar = RT_ColsAndRows Then
        Call SetVSFHead(mvsfBinding, mstrHead)
        Call SetColsHide
    End If
    If typVar = RT_Rows Or typVar = RT_ColsAndRows Then
        Call FillData
    End If
End Sub

Private Sub FillData()
    Dim intCol As Integer, intOldCol As Integer, intField As Integer
    Dim lngRow As Long, lngOldRow As Long, lngOldTop As Long
    
    With mvsfBinding
        intOldCol = .Col
        lngOldRow = .Row
        lngOldTop = .TopRow
        
        .Redraw = False
        .Clear 1
        .Rows = 1
                
        If mrsRecordset.State <> adStateOpen Then
            .Redraw = True
            'MsgBox "��¼��δ�򿪣�", vbInformation, App.ProductName
            Exit Sub
        End If
        
        If mrsRecordset.RecordCount > 0 Then mrsRecordset.MoveFirst
        Do While mrsRecordset.EOF = False
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            For intField = 0 To mrsRecordset.Fields.Count - 1
                intCol = .ColIndex(mrsRecordset.Fields(intField).name)
                If intCol >= 0 Then
                    .TextMatrix(lngRow, intCol) = IIf(IsNull(mrsRecordset.Fields(intField).value), "", mrsRecordset.Fields(intField).value)
                    Select Case mrsRecordset.Fields(intField).type
                    Case adDBTimeStamp
                        .TextMatrix(lngRow, intCol) = Format(.TextMatrix(lngRow, intCol), "YYYY-MM-DD hh:mm:ss")
                    Case adDBTime
                        .TextMatrix(lngRow, intCol) = Format(.TextMatrix(lngRow, intCol), "YYYY-MM-DD hh:mm:ss")
                    Case adDBDate, adDate
                        .TextMatrix(lngRow, intCol) = Format(.TextMatrix(lngRow, intCol), "YYYY-MM-DD")
                    End Select
                    RaiseEvent EventFillData(mvsfBinding, lngRow, intCol)
                End If
            Next
            
            '�и�
            If HeightRow > 0 Then
                .rowHeight(lngRow) = HeightRow
            End If
            
            mrsRecordset.MoveNext
        Loop
        
        '�ָ�����
        If mtypSort.Col >= 0 Then
            .Col = mtypSort.Col
            .ColSort(mtypSort.Col) = mtypSort.type
            .Sort = mtypSort.type
        End If
        
        '�ָ�����λ��
        If .Rows > 1 Then
            If lngOldRow > 0 Then
                If .Rows - 1 >= lngOldRow Then
                    .Row = lngOldRow
                Else
                    .Row = .Rows - 1
                End If
                
                '��ԭ�ϴε�TopRow
                If .Row >= .BottomRow And .Row <= .TopRow - 1 Then
                    '...
                Else
                    If .Row - (lngOldRow - lngOldTop) >= .Rows Then
                        .TopRow = .Rows - 1
                    ElseIf .Row - (lngOldRow - lngOldTop) <= 0 Then
                        .TopRow = 1
                    Else
                        .TopRow = .Row - (lngOldRow - lngOldTop)
                    End If
                End If
            Else
                .Row = 1
            End If
        End If
        .Col = intOldCol
        
        'ֻ���еı���ɫ
        Call SetColsBackColor
        
        .Redraw = True
    End With
End Sub

Private Sub Class_Initialize()
    mtypSort.Col = -1
End Sub

Private Sub Class_Terminate()
    Set mrsRecordset = Nothing
    Set mvsfBinding = Nothing
End Sub

Private Sub SetVSFHead(ByVal vsfObject As Object, ByVal strHead As String)
'--------------------------------
'���ܣ���ʼ��VSFlexGrid�ؼ�����ͷ
'������
'  vsfObject��Ŀ��ؼ���
'  strHead������ͷ�ĳ�ʼ���ִ�
'
'��ʽ�� "����,,3,1000,s|..."
'   Ԫ��1��Keyֵ��
'   Ԫ��2��Captionֵ��Ĭ��ΪKeyֵ����
'   Ԫ��3�������ԣ�0���ڲ���ʾ�����ƶ���1���ڲ����أ������ƶ���������ʾ��2���û����أ�3���û���ʾ(Ĭ��ֵ)��
'   Ԫ��4���п��ȣ�Ĭ��0����
'   Ԫ��5����ʾ��ʽ��s(Ĭ��)���ַ����� n�����֣� d�����ڣ� t��ʱ�䣻 dt������ʱ�䣻 b: �����ͣ� c��ѡ��򣨲����ͣ���
'--------------------------------
    Dim arrCols As Variant, arrRows As Variant
    Dim i As Integer
    
    On Error GoTo errHandle
    
    arrRows = Split(strHead, "|")
    With vsfObject
        If .Rows = 0 Then .Rows = 1
        .Cols = UBound(arrRows) + 1
        If Me.HeightColumn > 0 Then
            .rowHeight(0) = Me.HeightColumn
        End If
        For i = LBound(arrRows) To UBound(arrRows)
            If arrRows(i) <> "" Then
                arrCols = Split(arrRows(i), ",")
                '��1Ԫ�أ�Keyֵ
                .ColKey(i) = arrCols(0)
                
                '��2Ԫ�أ�Captionֵ
                If arrCols(1) = "" Then
                    .TextMatrix(0, i) = arrCols(0)
                Else
                    .TextMatrix(0, i) = arrCols(1)
                End If
                
                '��3Ԫ�أ�������
                If arrCols(2) = "" Then
                    .ColData(i) = 3
                Else
                    .ColData(i) = val(arrCols(2))
                End If
                
                '��4Ԫ�أ�����
                .ColWidth(i) = val(arrCols(3))
                
                '��5Ԫ�أ���ʾ��ʽ
                If UBound(arrCols) > 3 Then
                    If UCase(arrCols(4)) = "D" Then
                        .ColFormat(i) = "yyyy-mm-dd"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrCols(4)) = "T" Then
                        .ColFormat(i) = "hh:mm:ss"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrCols(4)) = "DT" Then
                        .ColFormat(i) = "yyyy-mm-dd hh:mm:ss"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrCols(4)) = "N" Then
                        .ColAlignment(i) = flexAlignRightCenter
                    Else
                        .ColAlignment(i) = flexAlignLeftCenter
                        If UCase(arrCols(4)) = "B" Then
                            .ColDataType(i) = flexDTBoolean
                        End If
                    End If
                Else
                    .ColAlignment(i) = flexAlignLeftCenter
                End If
                
                '������
                If val(arrCols(2)) = 1 Or val(arrCols(2)) = 2 Then
                    .ColHidden(i) = True
                Else
                    .ColHidden(i) = False
                End If
                
            End If
        Next
        
        If .Cols > 0 Then .Cell(2, 0, 0, 0, .Cols - 1) = 4
    End With
    Exit Sub
    
errHandle:
    MsgBox err.Description, vbInformation, App.ProductName
End Sub

Private Sub mvsfBinding_AfterSort(ByVal Col As Long, Order As Integer)
    Dim i As Long
    
    mtypSort.Col = Col
    mtypSort.type = Order
    
    If Order = 0 Then Exit Sub
    
    With mvsfBinding
        '�ָ�֮ǰ��ѡ����
        For i = 1 To .Rows - 1
            If .IsSelected(i) Then
                .Row = i
                .TopRow = i
                Exit For
            End If
        Next
    End With
End Sub

Private Sub mvsfBinding_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'ֻ����
    Cancel = InStr("|" & mstrColsReadonly & "|", "|" & mvsfBinding.ColKey(Col) & "|") > 0
End Sub

Private Sub mvsfBinding_GotFocus()
    '�õ����㵱ǰ�еı���ɫ
    mvsfBinding.BackColorSel = &H8000000D
End Sub

Private Sub mvsfBinding_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Shift = vbShiftMask Then
            PressKey vbKeyLeft
        ElseIf Shift = 0 Then
            PressKey vbKeyRight
            
            With mvsfBinding
                .Redraw = False
                If .Col >= .RightCol Then
                    '���һ��ʱ���н��㶨λ����һ��
                    If mbytEditable = EM_Modify And .Row >= .Rows - 1 Then
                        '���ݡ�AppendFields��������������
                        Call Append(mstrAppendFields)
                        .Row = .Rows - 1
                    Else
                        .Row = .Row + 1
                    End If
                    .ColSel = .LeftCol
                    
                    '�ػ��еı���ɫ�������������ӣ�
                    Call SetColsBackColor
                End If
                .Redraw = True
            End With
        End If
    ElseIf KeyCode = vbKeyDelete And mbytEditable = EM_Modify Then
        If mvsfBinding.Row <= 0 Then Exit Sub
        
        If MsgBox("ȷ��ɾ��������", vbQuestion + vbYesNo + vbDefaultButton2, GSTR_MSG) = vbYes Then
            If mvsfBinding.Rows > 2 Then
                mvsfBinding.RemoveItem mvsfBinding.Row
            Else
                mvsfBinding.Clear 1
            End If
            mvsfBinding.SetFocus
        End If
    End If
End Sub

Private Sub mvsfBinding_LostFocus()
    'ʧȥ���㵱ǰ�еı���ɫ
    mvsfBinding.BackColorSel = &H8000000C
End Sub

Public Sub Init(Optional ByVal blnPermitSort As Boolean = False)
'���ܣ���ʼ���ؼ��Ļ�������
'������
'  blnPermitSort��True���������ͷ����False������

    If mvsfBinding Is Nothing Then Exit Sub
    
    With mvsfBinding
        .Appearance = flexFlat
        .AllowUserResizing = flexResizeColumns
        If mblnAllowMultSelection Then
            .SelectionMode = flexSelectionListBox
            .AllowSelection = True
        Else
            .SelectionMode = flexSelectionByRow
            .AllowSelection = False
        End If
        .FixedCols = 0
        .FixedRows = 1
        .SheetBorder = .BackColor
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .ExplorerBar = IIf(blnPermitSort, flexExSortShowAndMove, flexExMove)
        .AutoResize = True
        .SheetBorder = &H8000000F
        .BackColorBkg = &H80000005
        .RowHeightMin = 270
        .FocusRect = flexFocusHeavy
        .Rows = 1
        .rowHeight(0) = 350
    End With
End Sub

Public Sub Append(Optional ByVal strCopyCols As String)
'���ܣ�ΪVSFlexGrid����һ��

    Dim i As Integer, intCol As Integer
    Dim cllSource As Collection
    Dim arrCopyCols As Variant
    
    arrCopyCols = Split(arrCopyCols, "|")
    
    With mvsfBinding
        If .Rows > 1 Then
            If IsEmpty(arrCopyCols) Then
                .Rows = .Rows + 1
                Exit Sub
            End If
            If UBound(arrCopyCols) >= 0 Then
                '�Ƚ����һ�е�ֵ�ŵ����ϱ���
                Set cllSource = New Collection
                For i = LBound(arrCopyCols) To UBound(arrCopyCols)
                    intCol = .ColIndex(arrCopyCols(i))
                    If intCol >= 0 Then
                        cllSource.Add .TextMatrix(.Rows - 1, intCol), arrCopyCols(i)  '"_" & intCol
                    End If
                Next
                '����һ�к��ٻ�д����
                .Rows = .Rows + 1
                For i = 1 To cllSource.Count
                    .TextMatrix(.Rows - 1, .ColIndex(arrCopyCols(i - 1))) = cllSource(i)
                Next
            Else
                .Rows = .Rows + 1
            End If
        Else
            .Rows = .Rows + 1
        End If
    End With
    
End Sub

Private Sub SetColsBackColor()
'���ܣ����ò��ɱ༭�еı���ɫ
    
    Dim intCol As Integer

    With mvsfBinding
        If .Rows <= 1 Then Exit Sub
    
        If mbytEditable = EM_Display Then
            .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = .BackColor
        ElseIf mbytEditable = EM_View Then
            .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = &H8000000F
        Else
            For intCol = 0 To .Cols - 1
                If InStr("|" & mstrColsReadonly & "|", "|" & .ColKey(intCol) & "|") > 0 Then
                    .Cell(flexcpBackColor, 1, intCol, .Rows - 1, intCol) = &H8000000F
                End If
            Next
        End If
    End With
End Sub

Public Sub SetColsHide()
'���ܣ�������
'������
'  strCols���������ַ�����������1[|��2[|...]]

    Dim arrCols As Variant
    Dim strCol As String
    Dim i As Integer, j As Integer
    Dim blnFind As Boolean
    
    If mvsfBinding Is Nothing Then Exit Sub
    
    With mvsfBinding
        For i = 0 To .Cols - 1
            blnFind = False
            arrCols = Split(mstrColsHide & "|", "|")
            For j = LBound(arrCols) To UBound(arrCols)
                strCol = UCase(Trim(arrCols(j)))
                If strCol = UCase(.ColKey(i)) Then
                    blnFind = True
                    Exit For
                End If
            Next
            .ColHidden(i) = blnFind
        Next
    End With
End Sub

Public Sub RepaintRowHeight()
'���ܣ��ػ��и�
'������
'  intHeight���и�ֵ
'  intColumnHeight���е��и�

    Dim l As Long

    With Binding
        .Redraw = False

        For l = 0 To .Rows - 1
            If l = 0 Then
                If Me.HeightColumn > 0 Then
                    .rowHeight(l) = Me.HeightColumn
                End If
            Else
                If Me.HeightRow > 0 Then
                    .rowHeight(l) = Me.HeightRow
                Else
                    Exit For
                End If
            End If
        Next

        .Redraw = True
    End With
End Sub

