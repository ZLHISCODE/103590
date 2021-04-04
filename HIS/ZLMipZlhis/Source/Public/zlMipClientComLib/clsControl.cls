VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'---------------------------------------------------------------------
'ģ�����ƣ�clsControl
'�������󣺲����ؼ��ĺ���
'�������򣺿ؼ�����+�������� + ��������
'˵��    ���ؼ�����������淶�е���д����������һ��Ϊ���ʣ���������Ϊ�ؼ����ԣ������漰����򲻺�ȷ���ģ������ñ�������Ĵʻ������
'��      ��LvwSortColumn��LvwFlatColumnHeader��FrmCheckInput
'--------------------------------------------------------------------------
Public Enum mTextAlign
    taLeftAlign = 0
    taCenterAlign = 1
    taRightAlign = 2
End Enum

Public Enum mTextType
    m�ı�ʽ = 0
    m����ʽ = 1
    m���ʽ = 2
    m�����ʽ = 3
End Enum

Public Sub CboSetIndex(ByVal hWnd_combo As Long, ByVal lngIndex As Long)
'���ܣ�����Combo�ؼ���Indexֵ
'Ϊһ��Combo�ؼ�ѡ���б�����ֲ�������Click�¼�
    Const CB_SETCURSEL = &H14E
    
    SendMessage hWnd_combo, CB_SETCURSEL, lngIndex, 0
End Sub

Public Sub CboSetWidth(ByVal hWnd_combo As Long, ByVal lngWidth As Long)
'���ܣ�����Combo�ؼ������б��Ŀ���
'�˴��Ŀ������������б��Ŀ��ȣ���������TWIPΪ��λ
    Const CB_SETDROPPEDWIDTH As Long = &H160

    SendMessage hWnd_combo, CB_SETDROPPEDWIDTH, lngWidth / Screen.TwipsPerPixelX, 0
End Sub

Public Sub CboSetHeight(cboControl As Object, ByVal lngHeight As Long)
'���ܣ�����Combo�ؼ��������б��ĸ߶�
'�˴��Ŀ������������б��ĸ߶ȣ���������TWIPΪ��λ
    SetWindowPos cboControl.hWnd, 0, 0, 0, cboControl.Width / Screen.TwipsPerPixelX, lngHeight / Screen.TwipsPerPixelY, SWP_NOMOVE
End Sub

Public Sub CboSetText(cboControl As Object, ByVal strText As String, Optional ByVal blnAfter As Boolean = True, Optional strSplit As String = "-")
'���ܣ������ı�������Combo�ؼ��ĵ�ǰֵ
'������cboControl  ׼�����õ�ComboBox�ؼ�
'      strText     ������ı���
'      blnAfter    ��ʾ�ڷָ���֮ǰ��֮��ȡֵ�����û�зָ�������ȡ֮��
'      strSplit    �ָ�����ͨ��Ϊ-
    Dim lngPos As Long
    Dim lngCount As Long
    Dim strTemp As String
    Dim blnMatch As Boolean
    
    For lngCount = 0 To cboControl.ListCount - 1
        strTemp = cboControl.List(lngCount)
        
        lngPos = InStr(strTemp, strSplit)
        If lngPos = 0 Then
            'ֱ�ӷ��������ַ���
            If strText = strTemp Then
                blnMatch = True
                Exit For
            End If
        Else
            If blnAfter = False Then
                'Բ��֮ǰ
                If strText = Mid(strTemp, 1, lngPos - 1) Then
                    blnMatch = True
                    Exit For
                End If
            Else
                If strText = Mid(strTemp, lngPos + 1) Then
                    blnMatch = True
                    Exit For
                End If
            End If
        End If
    Next
    If blnMatch = True Then
        '�Ѿ��ҵ�
        cboControl.ListIndex = lngCount
    Else
        If blnAfter = True Then
            '�����ʵ�����ݣ����Ϊǰ��ֻ�Ǳ���
            If strText <> "" Then
                cboControl.AddItem strText
                cboControl.ListIndex = cboControl.NewIndex
            End If
        End If
    End If
End Sub

Public Sub CboAppendText(cboControl As Object, KeyAscii As Integer)
'���ܣ���ComboBoxʵ������������Զ���ɵĹ���
'˵������Combox.KeyPress�¼��е���
    Dim strInput As String
    Dim lngStart As Long
    Dim lngIndex As Long
    Const CB_FINDSTRING = &H14C
    
    If cboControl.STYLE <> 0 Then Exit Sub
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    strInput = Chr(KeyAscii): KeyAscii = 0

    With cboControl
        '��¼�ϴεĲ����λ��
        lngStart = .SelStart + IIf(strInput <> "", 1, 0)

        '���ŵõ��û�������ɺ��ı����г��ֵ�����
        strInput = Mid(.Text, 1, .SelStart) & strInput & Mid(.Text, .SelStart + .SelLength + 1)

        '���ݼ�������ݵõ����ܵ��б���
        lngIndex = SendMessage(cboControl.hWnd, CB_FINDSTRING, -1, ByVal strInput)
        If lngIndex >= 0 Then
            '.ListIndex = lngIndex
            .Text = .List(lngIndex)
            .SelStart = Len(strInput)
            .SelLength = 100
        Else
            .Text = strInput
            .SelStart = lngStart
        End If
    End With
End Sub

Public Function CboLocate(ByVal cboObj As Object, ByVal strValue As String, Optional ByVal blnItem As Boolean = False) As Boolean
    'blnItem:True-��ʾ����ItemData��ֵ��λ������;False-��ʾ�����ı������ݶ�λ������
    Dim lngLocate As Long
    CboLocate = False
    For lngLocate = 0 To cboObj.ListCount - 1
        If blnItem Then
            If cboObj.ItemData(lngLocate) = Val(strValue) Then
                cboObj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        Else
            If Mid(cboObj.List(lngLocate), InStr(1, cboObj.List(lngLocate), "-") + 1) = strValue Then
                cboObj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        End If
    Next
End Function

Public Sub CboAddData(objSource As Object, ByVal rsDataSource As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
    '-----------------------------------------------------------------------------------
    '����:װ��������ָ�������������������е���������
    '����:objSource   Ҫװ�ؼ�¼����������ؼ�
    '     rsDataSource     ��¼������,ֻ��������������,��һ������������,�ڶ����Ƕ�Ӧ�Ĺؼ���,��id
    '     blnClear    װ��ʱ�Ƿ����ԭ�е���������,ȱʡΪTrue
    '����:
    '-----------------------------------------------------------------------------------
    If blnClear = True Then objSource.Clear
    
    If rsDataSource.BOF = False Then
        rsDataSource.MoveFirst
        While Not rsDataSource.EOF
            objSource.AddItem rsDataSource.Fields(0).value
            objSource.ItemData(objSource.NewIndex) = Val(rsDataSource.Fields(1).value)
            rsDataSource.MoveNext
        Wend
        rsDataSource.MoveFirst
    End If
End Sub

Public Function CboMatchIndex(ByVal lngHwnd As Long, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'���ܣ�����������ַ����Զ�ƥ��ComboBox��ѡ����,���Զ�ʶ��������
'������lngHwnd=ComboBox��Hwnd����,KeyAscii=ComboBox��KeyPress�¼��е�KeyAscii����,sngInterval=ָ��������
'���أ�ƥ�������(����ƥ�������)
'˵����ֻ��������,��δ�����������뽫�ú�����KeyPress�¼��е��á�
    Static lngPreTime As Single, lngPreHwnd As Long, strFind As String
    Dim sngTime As Single, lngR As Long
    Dim lngCount As Long, strText As String * 255
    Dim lngIdx As Long, strTmp As String, i As Long
    
    If lngHwnd <> 0 Then
        If KeyAscii = 13 Then '�س�������,ֱ�ӷ��ص�ǰ����
            CboMatchIndex = SendMessage(lngHwnd, CB_GETCURSEL, 0, 0)
            Exit Function
        Else
            '�����б�����
            If SendMessage(lngHwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 Then
                Call gobjComLib.zlCommFun.PressKey(vbKeyF4)
            End If
        End If
    End If
    
    If lngPreHwnd <> lngHwnd Or lngHwnd = 0 Then
        lngPreTime = Empty: strFind = Empty
    End If
    lngPreHwnd = lngHwnd
    
    If lngHwnd = 0 Then
        CboMatchIndex = -1: Exit Function
    End If
    
    sngTime = timer
    If Abs(sngTime - lngPreTime) > sngInterval Then '������(ȱʡΪ0.5��)
        strFind = ""
    End If
    strFind = strFind & Chr(KeyAscii)
    lngPreTime = timer
    KeyAscii = 0 'ʹComboBox�����ĵ���ƥ�书��ʧЧ
        
    '��API�ӿ�ͷ���в���
    lngIdx = SendMessage(lngHwnd, CB_FINDSTRING, -1, ByVal strFind)
    
    '��ƥ���ٴ��м䰴����ƥ��
    If lngIdx = -1 Then
        lngCount = SendMessage(lngHwnd, CB_GETCOUNT, 0, 0)
        For i = 0 To lngCount - 1
            strText = String(SendMessage(lngHwnd, CB_GETLBTEXTLEN, i, 0), Chr(32)) & Chr(0)
            Call SendMessage(lngHwnd, CB_GETLBTEXT, i, ByVal strText)
            
            If InStr(strText, Chr(0)) > 0 Then
                strTmp = Mid(strText, 1, InStr(strText, Chr(0)) - 1)
            Else
                strTmp = strText
            End If
            
            If InStr(strTmp, "]") > 0 And InStr(strTmp, "-") = 0 Then
                strTmp = LTrim(Mid(strTmp, InStr(strTmp, "]") + 1))
            ElseIf InStr(strTmp, ")") > 0 And InStr(strTmp, "-") = 0 Then
                strTmp = LTrim(Mid(strTmp, InStr(strTmp, ")") + 1))
            Else
                strTmp = LTrim(Mid(strTmp, InStr(strTmp, "-") + 1))
            End If
            
            If gobjComLib.zlCommFun.SpellCode(strTmp) Like UCase(strFind) & "*" Then
                lngIdx = i: Exit For
            End If
        Next
    End If
    
    If lngIdx = -1 Then Beep
    CboMatchIndex = lngIdx
End Function

Public Function FormCheckInput(objForm As Object) As Boolean
'���ܣ����ָ���������Ƿ��зǷ��ַ����뼰���볤���Ƿ񳬳�
'������objForm=Ҫ���Ĵ���
    Dim obj As Object, strText As String
    
    On Error Resume Next
    For Each obj In objForm.Controls
        If InStr("TextBox,ComboBox", TypeName(obj)) > 0 Then
            If obj.Visible And obj.Enabled And Not obj.Locked Then
                Select Case TypeName(obj)
                Case "TextBox"
                    strText = obj.Text
                Case "ComboBox"
                    If obj.STYLE = 0 Then strText = obj.Text
                End Select
                If InStr(strText, "'") > 0 _
                    Or InStr(strText, ",") > 0 _
                    Or InStr(strText, ";") > 0 _
                    Or InStr(strText, "|") > 0 _
                    Or InStr(strText, "~") > 0 _
                    Or InStr(strText, "^") > 0 Then
                    MsgBox "���������а����Ƿ��ַ���", vbInformation, gstrSysName
                    obj.SelStart = 0: obj.SelLength = Len(obj.Text)
                    obj.SetFocus: Exit Function
                End If
            End If
        End If
    Next
    FormCheckInput = True
End Function

'Public Function LvwSelectColumns(objSet As Object, ByVal strColumn As String, Optional ByVal blnInit As Boolean = False) As Boolean
''����:���б��ؼ����н�������
''����:
''   objSet��Ҫ���õĶ���,Ŀǰֻ֧��ListView���Ժ��ټ���FlexGrid,DataGrid��
''   strColumn���д�����ʽ��"����,�п�,������ֵ,������;����,�п�,������ֵ,������"    ע����֮�����÷ֺ�
''      ���� "����,2000,0,1;����,800,0,0;����,1440,0,0"
''      ��ListView���ԣ�������Ϊ1��ʾ���в���ɾ����������Ϊ0��ʾ���п���ɾ��
''      ��FlexGridView���ԣ������Ի�Ҫ��ʾ�Ƿ����ڹ̶��У��Ա㲻�ܺ������н���˳�����
''   blnInit��True,����ʾѡ�񴰿ڣ�ֱ�ӳ�ʼ��
'    Dim varColumns As Variant, varColumn As Variant
'    Dim lngCol As Long
'
'    If blnInit Then
'        varColumns = Split(strColumn, ";")
'        Select Case TypeName(objSet)
'            Case "ListView"
'                With objSet.ColumnHeaders
'                    .Clear
'                    For lngCol = LBound(varColumns) To UBound(varColumns)
'                        varColumn = Split(varColumns(lngCol), ",")
'                        .Add , "_" & varColumn(0), varColumn(0), varColumn(1), varColumn(2)
'                    Next
'                End With
'            Case "MSHFlexGrid"
'            Case "DataGrid"
'        End Select
'    Else
'        LvwSelectColumns = frmSelectColumn.������(objSet, strColumn)
'    End If
'End Function

Public Sub LvwFlatColumnHeader(ByVal lvw As Object)
'���ܣ�ʹ��ListView���б����Ϊƽ��
    Const strHeaderClass As String = "msvb_lib_header"
    Const HDS_BUTTONS   As Long = 2
    
    Dim lngChild As Long, lngLen As Long, lngStyle As Long
    Dim strName As String * 255

    
    lngChild = GetWindow(lvw.hWnd, GW_CHILD)
    Do While lngChild <> 0
        lngLen = GetClassName(lngChild, strName, 255)
        If lngLen > 0 Then
            If Mid(strName, 1, lngLen) = strHeaderClass Then
                lngStyle = GetWindowLong(lngChild, GWL_STYLE)
                lngStyle = lngStyle And (Not HDS_BUTTONS)
                SetWindowLong lngChild, GWL_STYLE, lngStyle
                Exit Sub
            End If
        End If
        lngChild = GetWindow(lngChild, GW_HWNDNEXT)
    Loop

End Sub

Public Sub LvwSortColumn(lvw As Object, lngColumn As Long)
'���ܣ�����ListView��ͷ�ĵ����������
'������lvw            ListView�ؼ�
'      lngColumn      ��ǰ�������ͷ��ţ���ColumnHeader.Indexֵ
    If lvw.Sorted = False Then lvw.Sorted = True

    If lvw.SortKey = lngColumn - 1 Then '���Ǹղ�����
        lvw.SortOrder = IIf(lvw.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvw.SortKey = lngColumn - 1  '����ؼ����Ǵ�0��ʼ��
        lvw.SortOrder = lvwAscending
    End If
    
    If Not lvw.SelectedItem Is Nothing Then
        lvw.SelectedItem.EnsureVisible
    End If
End Sub

Public Function LvwSaveItem(lvwObj As Object) As String
    '-----------------------------------------------------------------------------------
    '����:����ListView�е�ѡ����
    '����:lvwobj      Ҫ�����ListView�ؼ�
    '����:ѡ����Ĺؼ���
    '-----------------------------------------------------------------------------------
    If Not (lvwObj.SelectedItem Is Nothing) Then LvwSaveItem = lvwObj.SelectedItem.Key
End Function

Public Sub LvwRestoreItem(lvwObj As Object, svrKey As String)
    '-----------------------------------------------------------------------------------
    '����:�ָ�ListView�е�ѡ����
    '����:lvwobj      Ҫ�ָ���ListView�ؼ�
    '     svrKey      Ҫ�ָ���ListView��Ĺؼ���
    '����:
    '-----------------------------------------------------------------------------------
    On Error GoTo EndP
    If lvwObj.ListItems.count > 0 Then
        If Not (lvwObj.ListItems(svrKey) Is Nothing) Then
            lvwObj.ListItems(svrKey).Selected = True
            lvwObj.ListItems(svrKey).EnsureVisible
        End If
    End If
    Exit Sub
EndP:
    If lvwObj.ListItems.count > 0 Then lvwObj.ListItems(1).Selected = True
End Sub

Public Sub LvwSetColWidth(objLvw As Object, Optional blnHideNullCol As Boolean)
'���ܣ�����ListView�е�ǰ�������Զ�������Ϊ��Сƥ�����,���������ٿ�����ʾ��ͷ���ֵĿ���
'������objLvw=Ҫ������ListView����
'      blnHideNullCol=�Ƿ�����û���κ����ݵ���
    Dim i As Integer, lngW As Long
    For i = 1 To objLvw.ColumnHeaders.count
        SendMessage objLvw.hWnd, LVM_SETCOLUMNWIDTH, i - 1, LVSCW_AUTOSIZE
        If blnHideNullCol Then If objLvw.ColumnHeaders(i).Width < 200 Then objLvw.ColumnHeaders(i).Width = 0
        If objLvw.ColumnHeaders(i).Width < (gobjComLib.zlCommFun.ActualLen(objLvw.ColumnHeaders(i).Text) + 2) * 90 And objLvw.ColumnHeaders(i).Width <> 0 Then
            objLvw.ColumnHeaders(i).Width = (gobjComLib.zlCommFun.ActualLen(objLvw.ColumnHeaders(i).Text) + 2) * 90
        End If
    Next
End Sub

Public Sub MshSetColWidth(objMsh As Object, frmParent As Object)
'���ܣ��Զ�����MSHFlexGrid����ĸ��п���,����Сƥ��Ϊ׼,��������ͷ���ֿ���ȫ����ʾ
'������msh=Ҫ�����п���MSHFlexGrid�ؼ�
'      frmParent=������,���ڼ������ֿ���
    Dim arrWidth() As Long, blnRedraw As Boolean
    Dim lngRows As Long
    Dim i As Long, j As Long
    
    ReDim arrWidth(objMsh.Cols - 1)
    
    blnRedraw = objMsh.Redraw
    lngRows = objMsh.Rows - 1
    If lngRows > 200 Then lngRows = 200
    objMsh.Redraw = False
    For i = 0 To objMsh.Cols - 1
        For j = IIf(objMsh.FixedRows = 0, 0, objMsh.FixedRows - 1) To lngRows
            If Not objMsh.MergeRow(j) Then
                If frmParent.TextWidth(objMsh.TextMatrix(j, i)) + 90 > arrWidth(i) Then
                    arrWidth(i) = frmParent.TextWidth(objMsh.TextMatrix(j, i)) + 90
                End If
            End If
        Next
    Next
    For i = 0 To objMsh.Cols - 1
        objMsh.ColWidth(i) = IIf(arrWidth(i) > 3000, 3000, arrWidth(i))
    Next
    objMsh.Redraw = blnRedraw
End Sub

Public Sub MshColumnSort(ByVal mshGrd As Object, ByVal IntSorCol As Integer, _
    ByRef intPreCol As Integer, ByRef intPreSort As Integer, _
    Optional blnNum As Boolean = False, Optional intMachingCol As Integer = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:���н�������
    '����:
    '       mshGrd-MSHFlexGrid����ؼ�
    '       intPreCol-ǰһ�ε���ͷ��������
    '       intPreSort-ǰһ�ε���ͷ������
    '       blnNum-�Ƿ�Ϊ������
    '       intMachingCol-ƥ����,��Ҫ������������¶�λ��ָ������
    '---------------------------------------------------------------------------------------------------------------------------------------------------------
    Const flexSortNumericAscending  As Long = 3
    Const flexSortNumericDescending  As Long = 4
    Const flexSortStringNoCaseAscending As Long = 5
    Const flexSortStringNoCaseDescending As Long = 6
    
    Dim intCol As Integer
    Dim intRow As Integer
    Dim strTmp As String
    
    Err = 0
    On Error GoTo errHand:
    With mshGrd
        If IntSorCol > .Cols - 1 Or IntSorCol < 0 Then Exit Sub
        If .Rows > 1 Then
            .Redraw = False
            intCol = IntSorCol   '.MouseCol
            .Col = intCol
            .ColSel = intCol
            If intMachingCol > .Cols - 1 Or intMachingCol < 0 Then
                strTmp = .TextMatrix(.Row, 0)
            Else
                strTmp = .TextMatrix(.Row, intMachingCol)
            End If
            If blnNum Then
                    If intCol = intPreCol And intPreSort = flexSortNumericDescending Then
                       .Sort = flexSortNumericAscending
                       intPreSort = flexSortNumericAscending
                    Else
                       .Sort = flexSortNumericDescending
                       intPreSort = flexSortNumericDescending
                    End If
            Else
                    If intCol = intPreCol And intPreSort = flexSortStringNoCaseDescending Then
                       .Sort = flexSortStringNoCaseAscending
                       intPreSort = flexSortStringNoCaseAscending
                    Else
                       .Sort = flexSortStringNoCaseDescending
                       intPreSort = flexSortStringNoCaseDescending
                    End If
            End If
            
            intPreCol = intCol
            If intMachingCol > .Cols - 1 Or intMachingCol < 0 Then
                .Row = MshGrdFindRow(mshGrd, strTmp, 0)
            Else
                .Row = MshGrdFindRow(mshGrd, strTmp, intMachingCol)
            End If
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
    Exit Sub
errHand:
    mshGrd.Redraw = True
End Sub

Public Function MshGrdFindRow(ByVal mshGrd As Object, _
        ByVal strTmp As Variant, ByVal intCol As Integer) As Integer
    '----------------------------------------------------------------------------
    '����:Ѱ����ĳһ��Ԫֵ��ȵ���
    '����:
    '       mshGrd-����ؼ�
    '       strTmp-ƥ�䴮
    '       intCol-ƥ����
    '����:���ҵ�����,��δ�ҵ�,��λ��һ��
    '----------------------------------------------------------------------------
    
    Dim i As Integer
    Err = 0
    On Error GoTo errHand:
    
    With mshGrd
        For i = 1 To .Rows - 1
            If IsDate(strTmp) Then
               If Format(.TextMatrix(i, intCol), "yyyy-mm-dd") = Format(strTmp, "yyyy-mm-dd") Then
                  MshGrdFindRow = i
                  Exit Function
               End If
            Else
                If .TextMatrix(i, intCol) = strTmp Then
                  MshGrdFindRow = i
                  Exit Function
                End If
            End If
        Next
    End With
    MshGrdFindRow = 1
    Exit Function
errHand:
    MshGrdFindRow = 1
    Exit Function
End Function

Public Sub MshClearData(obj As Object, ByVal intRow As Integer, Optional intCol As Variant)
    '-----------------------------------------------------------------------------------
    '����: ���ָ�������ָ����ָ���е�����
    '����: obj=Ҫ����������ؼ�
    '      intRow=Ҫ������к�
    '      intCol=Ҫ������к��б���Array(1,2,3),������������Ա�ʾΪArray()
    '���أ�
    '-----------------------------------------------------------------------------------
    Dim i As Long
    If UBound(intCol) = -1 Then
        For i = 0 To obj.Cols - 1
            obj.TextMatrix(intRow, i) = ""
        Next
    Else
        For i = 0 To UBound(intCol)
            obj.TextMatrix(intRow, intCol(i)) = ""
        Next
    End If
    obj.RowData(intRow) = 0
End Sub

Public Sub MshSetFormat(ByVal msfObj As Object, ByVal strColumns As String, ByVal strCaption As String, _
Optional ByVal rsData As ADODB.Recordset, Optional ByVal blnData As Boolean = False, Optional ByVal blnRestore As Boolean = False)
    Dim intCol As Integer, intCols As Integer
    '���������壺msfObj�Ǳ������strColumns��������ʽ��������|����|�ж��뷽ʽ����blnData˵�������ø�ʽ���ǰ����ݣ�blnRestore˵���Ƿ�ָ�ԭʼ����
    '������ʽ��NO|1000|4,����|1600|4,��鵥λ|2000|4,�����|800|4,��Ҫ��¼|2000|4,����|1000|4
    
    '���ñ���ؼ���ʽ
    intCols = UBound(Split(strColumns, ",")) + 1
    msfObj.Redraw = False
    
    '���õ���ͷ��������ĸ�ʽ
    If Not blnData Then
        With msfObj
            .Clear
            .Rows = 2
            .Cols = intCols
            
            For intCol = 0 To .Cols - 1
                .TextMatrix(0, intCol) = Split(Split(strColumns, ",")(intCol), "|")(0)
                .ColAlignmentFixed(intCol) = 4
            Next
            If blnRestore Then
                For intCol = 0 To .Cols - 1
                    .ColWidth(intCol) = Split(Split(strColumns, ",")(intCol), "|")(1)
                Next
                Call gobjComLib.RestoreFlexState(msfObj, strCaption)
            End If
        End With
    Else
        If Not rsData.EOF Then
            With msfObj
                Set .DataSource = rsData
                For intCol = 0 To .Cols - 1
                    .ColAlignment(intCol) = Split(Split(strColumns, ",")(intCol), "|")(2)
                Next
            End With
        End If
    End If
    
    With msfObj
        .Row = 1: .Col = 0
        .RowSel = 1: .ColSel = .Cols - 1
        .Redraw = True
    End With
End Sub

Public Sub VSFSetFontSize(ByRef vsf As Object, ByVal lngFontSize As Long, Optional ByVal lngCol As Long)
'���ܣ�����vsfflexgrid����Ĵ�С�����Զ������п����и�
'������lngFontSize�������ã�9�ż�С�壬12�ż�С��
'      lngCol,���Ҫ�����Զ������иߣ������и����ݵ��У�Ҫ��AutoSizeMode = flexAutoSizeRowHeight,WordWrap =True
    Dim i As Long, lngRate As Double, lngTmp As Long
    
    If lngFontSize < 5 And lngFontSize > 50 Then Exit Sub
    With vsf
        lngRate = lngFontSize / .FontSize
        lngTmp = .Redraw
        
        .Redraw = flexRDNone
        .FontSize = lngFontSize
        
        For i = 0 To .Cols - 1
           If .ColWidth(i) > 0 Then
             .ColWidth(i) = .ColWidth(i) * lngRate
           End If
        Next
        
        If .AutoSizeMode = flexAutoSizeRowHeight And .WordWrap And lngCol > 0 Then
            .AutoSize lngCol
        Else
            .RowHeightMin = .RowHeightMin * lngRate
            .RowHeight(-1) = .RowHeightMin
        End If
        .Redraw = lngTmp
    End With
End Sub

Public Sub PicShowFlat(objPic As Object, Optional intStyle As Integer = -1, Optional strName As String = "", Optional intAlign As mTextAlign)
'���ܣ���PictureBoxģ��ɰ��»�͹������
'������intStyle:0=ƽ��,-1=����,1=͹��
'      intAlign=���Ҫ��ʾ�ı�,��ָ�����뷽ʽ
    
    Dim vRect As RECT, lngTmp As Long
    
    With objPic
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            vRect.Left = .ScaleLeft
            vRect.Top = .ScaleTop
            vRect.Right = .ScaleWidth
            vRect.Bottom = .ScaleHeight
            DrawEdge .hDC, vRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            If intAlign = taCenterAlign Then
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2 '�м����
            ElseIf intAlign = taRightAlign Then
                .CurrentX = .ScaleWidth - .TextWidth(strName) - 2 '�ұ߶���
            Else
                .CurrentX = 2 '��߶���
            End If
            objPic.Print strName
        End If
    End With
End Sub

Public Sub StaShowPercent(sngPercent As Single, objPanel As Object, objForm As Object)
'���ܣ���״̬���ϸ��ݰٷֱ���ʾ��ǰ��������(��)
'������sngPercent=��ǰ����ֵ,Ϊ0-1֮���һ��С��
'      objPanel=״̬����������ʾ���ȵĸ���
'      objForm=���ڼ������ֿ��ȵĴ���
    Dim intAll As Integer
    intAll = objPanel.Width / objForm.TextWidth("��") - 4
    objPanel.Text = Format(sngPercent, "0% ") & String(intAll * sngPercent, "��")
End Sub

Public Sub TxtGetCaretPos(ByVal lngTxtHwnd As Long, lngLineNo As Long, lngColNo As Long)
'���ܣ���ȡ������ı����е�����λ��
'������lngTxtHwnd=Ҫ��ȡ���λ�õ��ı�����
'���أ�lngLineNo=��������к�,��1��ʼ
'      lngColNo=��������к�,��0��ʼ
    Dim i As Long, j As Long, K As Long
    Dim lParam As Long, wParam As Long

    '�������ı��򴫵�EM_GETSEL��Ϣ�Ի�ȡ����ʼλ�õ�
    '�������λ�õ��ַ���
    i = SendMessage(lngTxtHwnd, EM_GETSEL, wParam, lParam)
    j = i / 2 ^ 16
    
    '�����ı��򴫵�EM_LINEFROMCHAR��Ϣ���ݻ�õ��ַ�
    '��ȷ������Ի�ȡ��������
    lngLineNo = SendMessage(lngTxtHwnd, EM_LINEFROMCHAR, j, 0) '
    lngLineNo = lngLineNo + 1
    
    '���ı��򴫵�EM_LINEINDEX��Ϣ�Ի�ȡ��������
    K = SendMessage(lngTxtHwnd, EM_LINEINDEX, -1, 0)
    lngColNo = j - K + 1
End Sub

Public Sub TxtGetCursorPos(objMain As Object, objTxt As Object, lngX As Long, lngY As Long)
'���ܣ���ȡ�ı����е�ǰ�������ڸÿ������
'������objMain=���ڼ������ֳߴ�Ĵ���
'      objTXT=Ҫ��ȡ���������ı���ؼ�
'���أ� lngX,lngY=�������(Twip)
    Dim lngRow As Long, lngCol As Long
    Dim lngFirst As Long
    Dim lngRH As Long, lngCW As Long
    
    objMain.Font.Name = objTxt.Font.Name
    objMain.Font.Size = objTxt.Font.Size
    objMain.Font.Bold = objTxt.Font.Bold
    objMain.Font.Italic = objTxt.Font.Italic
    objMain.Font.Underline = objTxt.Font.Underline
    objMain.Font.Strikethrough = objTxt.Font.Strikethrough
    lngRH = objMain.TextHeight("A")
    lngCW = objMain.TextWidth("A") '���԰�Ǽ����к�
    
    lngFirst = SendMessage(objTxt.hWnd, EM_GETFIRSTVISIBLELINE, lngRow, lngCol) + 1 '��0�п�ʼ
    Call TxtGetCaretPos(objTxt.hWnd, lngRow, lngCol) '��1��,0�п�ʼ
    lngRow = lngRow - lngFirst + 1 '��ƫ��
    
    lngY = lngRow * lngRH
    lngX = lngCol * lngCW
End Sub

Public Sub TxtCheckKeyPress(ByVal objCtl As Object, KeyAscii As Integer, ByVal TextType As mTextType)
    '------------------------------------------------------------------------------------------------------------------
    '����:ֻ���������ֺͻس����˸�
    '����:
    '   objctl:���ƿؼ�
    '   Keyascii:
    '           Keyascii:8 (�˸�)
    '   TextType:(0-�ı�ʽ;1-����ʽ;2-���ʽ)
    '����:һ��KeyAscii
    '------------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo errHand:
    
    If TextType = m�ı�ʽ Then
        If KeyAscii = Asc("'") Then
            KeyAscii = 0
        End If
        Exit Sub
    End If

    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        Select Case KeyAscii
        Case vbKeyReturn       '�س�
        Case 8                 '�˸�
        Case Asc(".")
            If TextType = m���ʽ Or TextType = m�����ʽ Then
                If InStr(objCtl, ".") <> 0 Then     'ֻ�ܴ���һ��С����
                    KeyAscii = 0
                End If
            Else
                KeyAscii = 0
            End If
        Case Asc("-")          '����
            Dim iRow As Long
            Dim iCol As Long
            If Trim(objCtl.Text) = "" Then Exit Sub
            If TextType <> m�����ʽ Then KeyAscii = 0: Exit Sub
            If objCtl.SelStart <> 0 Then KeyAscii = 0: Exit Sub      '��겻���һλ,�������븺��
            If InStr(1, objCtl.Text, "-") <> 0 Then   'ֻ�ܴ���һ������
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0
        End Select
    End If
    Exit Sub
errHand:
    KeyAscii = 0
End Sub

Public Sub TxtSelAll(objTxt As Object)
'���ܣ����༭��ĵ��ı�ȫ��ѡ��
'������objTxt=��Ҫȫѡ�ı༭�ؼ�,�ÿؼ�����SelStart,SelLength����
    objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    If TypeName(objTxt) = "TextBox" Then
        If objTxt.MultiLine Then
            SendMessage objTxt.hWnd, WM_VSCROLL, SB_TOP, 0
        End If
    End If
End Sub

Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'���ܣ���ʾ������һ������ı�����
'������blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(objForm.hWnd, vRect)
    lngStyle = GetWindowLong(objForm.hWnd, GWL_STYLE)
    If blnCaption Then
        lngStyle = lngStyle Or WS_CAPTION Or WS_THICKFRAME
        If objForm.ControlBox Then lngStyle = lngStyle Or WS_SYSMENU
        If objForm.MaxButton Then lngStyle = lngStyle Or WS_MAXIMIZEBOX
        If objForm.MinButton Then lngStyle = lngStyle Or WS_MINIMIZEBOX
    Else
        If blnBorder Then
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
        Else
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
        End If
    End If
    SetWindowLong objForm.hWnd, GWL_STYLE, lngStyle
    SetWindowPos objForm.hWnd, 0, vRect.Left, vRect.Top, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub