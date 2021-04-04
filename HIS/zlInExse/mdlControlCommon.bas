Attribute VB_Name = "mdlControlCommon"

'String
Public Function StringsDelItem(ByVal strItem As String, ByVal strList As String, Optional ByVal strSplit As String = ",") As String
'����:�Ƴ��ַ����е�һ��
'����:blnSplit-�ָ���
    Dim i As Long, arrTmp As Variant
    
    If strList = "" Then Exit Function
    
    arrTmp = Split(strList, strSplit)
    For i = 0 To UBound(arrTmp)
        If arrTmp(i) <> strItem Then
            StringsDelItem = StringsDelItem & "," & arrTmp(i)
        End If
    Next
    StringsDelItem = Mid(StringsDelItem, 2)
End Function

'ReportControl
Public Function GetRptColumn(ByRef rpt As ReportControl, ByVal strColumn As String) As Long
'���ܣ�����������������ţ�û���ҵ�ʱ����-1
    Dim arrTmp As Variant, i As Long
    
    GetRptColumn = -1
    For i = 0 To rpt.Columns.Count - 1
        If rpt.Columns(i).Caption = strColumn Then GetRptColumn = i: Exit For
    Next
End Function

Public Sub rptRemoveGroupsItem(ByRef rpt As ReportControl, ByVal strColumn As String)
'����:�Ƴ������е�һ��������,���·���
'����:rco-�����м���
    Dim i As Long, arrTmp As Variant
    Dim strGroup As String
    
    With rpt
        .Columns(GetRptColumn(rpt, strColumn)).Visible = True
        For i = 0 To .GroupsOrder.Count - 1
            If .GroupsOrder.Column(i).Caption <> strColumn Then
                strGroup = strGroup & IIf(strGroup = "", "", ",") & .GroupsOrder.Column(i).Index
            End If
        Next
        
        .GroupsOrder.DeleteAll
        arrTmp = Split(strGroup, ",")
        For i = 0 To UBound(arrTmp)
            .GroupsOrder.Add .Columns(arrTmp(i))
        Next
        .Populate
    End With
End Sub

'ComboBox
Public Sub CboAddByStrings(cbo As ComboBox, strList As String, Optional blnSelectOne As Boolean = True)
'����:�����ַ���ֵ��ʼ��Combobox�ؼ�
'����:strTmp-�Զ��ŷָ����ַ���,blnSelectOne-�Ƿ�ȱʡѡ�е�һ��
    Dim i As Long, arrTmp As Variant
    
    If strList = "" Then Exit Sub
    
    cbo.Clear
    arrTmp = Split(strList, ",")
    For i = 0 To UBound(arrTmp)
        cbo.AddItem arrTmp(i)
    Next
    If blnSelectOne Then cbo.ListIndex = 0
End Sub

Public Sub CboRemoveItem(cbo As ComboBox, strListName As String, Optional blnSelectOne As Boolean = True)
'����:�����ַ���ɾ��Combobox�ؼ��е���
'����:strTmp-Ҫɾ�����������,blnSelectOne-���ɾ����ǰ��ʾ����,��ɾ����Listindex=-1,�Ƿ�ȱʡѡ�е�һ��
    Dim i As Long, lngRow As Long
    
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = strListName Then
            cbo.RemoveItem (i)
            Exit For
        End If
    Next
    If cbo.ListIndex = -1 And blnSelectOne Then cbo.ListIndex = 0
End Sub



'vsfFlexGrid
Public Function VsfGetColNum(vsf As VSFlexGrid, strColName As String) As Long
'����:������������vsfFlexGrid�ؼ��е������,û���ҵ�ʱ����-1(ʹ��vsfFee.ColIndex������Ч)
'����:strColName-����
    Dim i As Long
    
    For i = 0 To vsf.Cols - 1
        If vsf.TextMatrix(0, i) = strColName Then VsfGetColNum = i: Exit Function
    Next
    VsfGetColNum = -1
End Function

'MSHFlexGrid
Public Function MshGetColNum(msh As MSHFlexGrid, strColName As String) As Long
'����:������������MSHFlexGrid�ؼ��е������,û���ҵ�ʱ����-1
'����:strColName-����
    Dim i As Long
    
    For i = 0 To msh.Cols - 1
        If msh.TextMatrix(0, i) = strColName Then MshGetColNum = i: Exit Function
    Next
    MshGetColNum = -1
End Function


Public Function MshLackComplement(msh As MSHFlexGrid, txtLack As TextBox, ByVal strDec As String, ByVal curMustSum As Currency, _
    ByVal lngCurrencyCol As Long, ByVal lngDefaultRow As Long, ByVal lngCurrentRow As Long, ByVal blnDefaultRowIsCash As Boolean, _
    Optional ByVal lngStartRow As Long = 1) As Currency
'����:���ݵ�ǰ�����Ѹı�ĵ�Ԫ���ֵ,��"Ӧ���ܽ��"-"�Ѹ��ܽ��"�Ĳ��ӵ�ȱʡ����
'     �����ǰ����ȱʡ��,������ʾ������ı�����
'����:CentMoney-�ֱҴ�����,gBytMoney-�ֱҴ������
'����:strDec-����ı�����ʾ�Ľ��С����ʽ(��:0.00),curMustSum-Ӧ���ܽ��
'     lngCurrencyCol-�����,lngDefaultRow-ȱʡ��,lngCurrentRow-��ǰ��,blnDefaultRowIsCash-ȱʡ���Ƿ���зֱҴ���,lngStartRow-��ʼ������(����������)
'����:�����
    Dim i As Long, curSum As Currency, curLack As Currency, curTmp As Currency, curError As Currency
        
    For i = lngStartRow To msh.Rows - 1
        curSum = curSum + Val(msh.TextMatrix(i, lngCurrencyCol))
    Next
    curLack = curMustSum - curSum
    
    If lngCurrentRow <> lngDefaultRow Then
        curTmp = Val(msh.TextMatrix(lngDefaultRow, lngCurrencyCol)) + curLack
        If curTmp = 0 Then
            msh.TextMatrix(lngDefaultRow, lngCurrencyCol) = ""
        Else
            If blnDefaultRowIsCash Then
                msh.TextMatrix(lngDefaultRow, lngCurrencyCol) = Format(CentMoney(curTmp), "0.00")
            Else
                msh.TextMatrix(lngDefaultRow, lngCurrencyCol) = Format(curTmp, "0.00")
            End If
        End If
        curLack = curTmp - Val(msh.TextMatrix(lngDefaultRow, lngCurrencyCol))
        curError = -1 * curLack
        txtLack.Text = Format(0, strDec)
    Else
        curError = 0
        txtLack.Text = Format(curLack, strDec)
        If curLack <> 0 And (Abs(curLack) < 0.1 Or gBytMoney = 5 And Abs(curLack) < 0.3) Then
            '1.�������ֽ�ֱҴ�����������,��������������ʱ��������0.29�����,0.79��0.5,0.29��0
            If blnDefaultRowIsCash Then
                curTmp = Val(msh.TextMatrix(lngDefaultRow, lngCurrencyCol))
                If CentMoney(curTmp + curLack) = curTmp Then txtLack.Text = Format(0, strDec): curError = -curLack
            Else
            '2.���ܽ��С��λ��ȡ����������
                If Abs(curLack) < 0.005 Then txtLack.Text = Format(0, strDec): curError = -curLack
            End If
        End If
    End If
    
    MshComplement = curError
End Function
