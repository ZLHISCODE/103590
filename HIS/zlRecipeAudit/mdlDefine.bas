Attribute VB_Name = "mdlDefine"
Option Explicit

'ģ���
Public Enum enuModule
    ���ﴦ�����_1351 = 1351
    סԺҩ�����_1352 = 1352
    ���������Ŀ_1353 = 1353
    �����������_1354 = 1354
    �������ͳ��_1355 = 1355
End Enum

Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Enum enuMenus
    �ļ� = 1
        ��ӡ���� = 101
        ��ӡԤ�� = 102
        ��ӡ = 103
        ���Excel = 104
        �������� = 181
        �˳� = 191
    �༭ = 2
        ������� = 8401
        ֹͣ��� = 8402
        �ϸ� = 3950
        ���ϸ� = 3951
    ���� = 4
    �鿴 = 5
        ������ = 701
            ��׼��ť = 7011
            �ı���ǩ = 7012
            ��ͼ�� = 7013
        ״̬�� = 702
        �����С = 509
            С���� = 4041
            ������ = 4042
        ˢ�� = 791
        �鿴PASS��� = 3944
    ���� = 6
        �������� = 901
        WEB�ϵ����� = 902
            ������ҳ = 9021
            ������̳ = 9023
            ���ͷ��� = 9022
        ���� = 991
End Enum

Public gobjPubAdvice As zlPublicAdvice.clsPublicAdvice     '�ٴ���������
Public gcnOracle As ADODB.Connection
Public gcnBusiness As ADODB.Connection
Public gstrSQL As String
Public glngSys As Long
Public glngModule As Long
Public gstrUnitName As String
Public gstrSysName As String                'ϵͳ����
Public gstrProductName As String            'OEM��Ʒ����
Public gintHoursRecipe As Integer           '����飬�ο�����Сʱ�ڵĴ���ҩƷ
Public gstrErrInfo As String

Public Sub AddArray(ByRef cllData As Collection, ByVal strSQL As String)
'���ܣ���SQLд�뼯��
'������
'  cllData�����϶���
'  strSQL��SQL�ַ���

    Dim l As Long
    
    l = cllData.Count + l
    cllData.Add strSQL, "K" & l
End Sub

Public Sub ExecuteProcedureArray(ByVal varArr As Variant, ByVal strCaption As String, Optional blnNoTrans As Boolean = False)
'����:ִ�ж����洢����
'����:
'  varArr��SQL���϶���
'  strCaption���������
'  blnNoTrans���Ƿ񲻴�������

    Dim i As Long, strSQL As String
    
    If blnNoTrans = False Then gcnOracle.BeginTrans
    For i = 1 To varArr.Count
        strSQL = varArr(i)
        zlDatabase.ExecuteProcedure strSQL, strCaption
    Next
    
    If blnNoTrans = False Then gcnOracle.CommitTrans
End Sub

'Public Function FormatString(ByVal strFormat As String, ParamArray arrParams() As Variant) As String
''���ܣ���ʽ���ַ���
''������
''  strFormat�����ʽ��[1-x]Ϊ�����Źؼ��֣����ӣ�"����ֵΪ��[1]"
''  arrParams�����ʽ�Ĳ�������ӦstrFormat�еĲ����Źؼ���
''���أ���ʽ������ַ���
'
'    Dim i As Integer, intSN As Integer
'    Dim strKey As String, strTmp As String
'    Dim blnStart As Boolean
'
'    FormatString = strFormat
'
'    If Len(strFormat) > 60000 Then Exit Function
'    If Not strFormat Like "*[[]*[]]*" Then Exit Function
'    If UBound(arrParams) < 0 Then Exit Function
'
'    On Error GoTo errHandle
'
'    For i = 1 To Len(strFormat)
'        If Mid(strFormat, i, 1) = "[" Then
'            blnStart = True
'        End If
'        If blnStart Then
'            If Mid(strFormat, i, 1) = "]" Then
'                intSN = Val(Mid(strKey, 2))
'                If intSN > 0 Then
'                    If UBound(arrParams) >= intSN - 1 Then
'                        strTmp = strTmp & arrParams(intSN - 1)
'                    End If
'                Else
'                    strTmp = strTmp & Mid(strKey, 2)
'                End If
'                blnStart = False
'                strKey = ""
'            Else
'                strKey = strKey & Mid(strFormat, i, 1)
'            End If
'        Else
'            strTmp = strTmp & Mid(strFormat, i, 1)
'        End If
'    Next
'
'    FormatString = strTmp
'    Exit Function
'
'errHandle:
'End Function

Public Sub SetLVColumnHeaders(ByRef lvwVar As ListView, ByVal strHeader As String)
'���ܣ�ͳһ����ListView����ͷ
'������
'  lvwVar��Ҫ���õ�ListView�ؼ�
'  strHeader����ͷ��׼�ִ�
'    ��ʽ������,Keyֵ,���,���뷽ʽ,ͼ���[|����1,...]
'    ˵����Keyֵ�����ʾ������������Ȳ����ʾ���أ����벻�Ĭ�����룻

    Dim i As Integer, j As Integer
    Dim arrCols As Variant, arrElements As Variant
    Dim strText As String, strKey As String
    Dim intWidth As Integer, intAlignment As Integer, intIcon As Integer

    If Trim(strHeader) = "" Then Exit Sub
    If lvwVar Is Nothing Then Exit Sub
    
    arrCols = Split(strHeader, "|")
    With lvwVar
        .ColumnHeaders.Clear
        
        For i = LBound(arrCols) To UBound(arrCols)
            arrElements = Split(arrCols(i), ",")
            If UBound(arrElements) < 2 Then
                MsgBox zlStr.FormatString("���á�[1]���ؼ���ͷ�Ĳ�������ȷ��", lvwVar.Name), vbInformation, gstrSysName
                Exit Sub
            End If
            '����
            strText = Trim(arrElements(0))
            If strText = "" Then
                MsgBox zlStr.FormatString("���á�[1]���ؼ���ͷ���ƵĲ�������ȷ��", lvwVar.Name), vbInformation, gstrSysName
                Exit Sub
            End If
            'Key
            If UBound(arrElements) > 0 Then
                strKey = arrElements(1)
            End If
            If Trim(strKey) = "" Then
                strKey = strText
            End If
            '���
            If UBound(arrElements) > 1 Then
                intWidth = Val(arrElements(2))
            Else
                intWidth = 0
            End If
            '����
            If UBound(arrElements) > 2 Then
                intAlignment = Val(arrElements(3))
                If intAlignment > 2 Then intAlignment = 0
            Else
                intAlignment = 0
            End If
            'ͼ���
            If UBound(arrElements) > 3 Then
                intIcon = Val(arrElements(4))
            Else
                intIcon = 0
            End If
            
            .ColumnHeaders.Add i + 1, strText, strKey, intWidth, intAlignment, intIcon
        Next
    End With
    
End Sub

Public Function GetLVColumnIndex(ByVal lvwVar As ListView, ByVal strKey As String) As Integer
'���ܣ���ȡָ��ListView�ؼ��е�Index
'������
'  lvwVar��ָ��ListView�ؼ�
'  strKey��Ҫ��ȡ�е�Keyֵ
'���أ��е�Index

    Dim i As Integer

    With lvwVar
        For i = 0 To .ColumnHeaders.Count - 1
            If UCase(strKey) = UCase(.ColumnHeaders.Item(i).Key) Then
                GetLVColumnIndex = i
                Exit Function
            End If
        Next
    End With

    GetLVColumnIndex = -1
End Function

Public Sub FillLVData(ByRef rsVar As ADODB.Recordset, ByRef lvwVar As ListView, _
    Optional ByVal strCheckCol As String, _
    Optional ByVal strKey As String)
'���ܣ���ListView�ؼ��������
'������
'  rsVar����¼������
'  lvwVar��ָ��Ҫ����ListView�Ŀؼ�
'  strCheckCol��Checkbox����
'  strKey��ָ����¼��ΪKey���ֶ�

    If rsVar Is Nothing Then Exit Sub
    If rsVar.State <> adStateOpen Then Exit Sub
    If lvwVar Is Nothing Then Exit Sub
    
    Dim limTmp As ListItem
    Dim i As Integer, j As Integer
    Dim arrFields As Variant
    Dim strMasterCol As String, strTmp As String
    
    strCheckCol = "," & strCheckCol & ","
    
    '����
    strMasterCol = lvwVar.ColumnHeaders(1).Key
    
    '������
    i = 1
    If rsVar.RecordCount > 0 Then rsVar.MoveFirst
    Do While rsVar.EOF = False
        '����
        strTmp = rsVar.Fields(strMasterCol).Value
        If InStr(strCheckCol, "," & strMasterCol & ",") > 0 Then
            '����ʾ�ı�
            strTmp = ""
        End If
        If strKey = "" Then
            Set limTmp = lvwVar.ListItems.Add(i, , strTmp)
        Else
            Set limTmp = lvwVar.ListItems.Add(i, "_" & rsVar.Fields(strKey).Value, strTmp)
        End If
        If lvwVar.Checkboxes Then
            limTmp.Checked = zlCommFun.NVL(rsVar.Fields(strMasterCol).Value, 0) = 1
        End If
        '����
        For j = 1 To lvwVar.ColumnHeaders.Count
            If j > 1 Then
                strTmp = lvwVar.ColumnHeaders(j).Key
                On Error Resume Next
                limTmp.ListSubItems.Add , , rsVar.Fields(strTmp).Value
                Err.Clear
                On Error GoTo 0
            End If
        Next
    
        rsVar.MoveNext: i = i + 1
    Loop

End Sub

Public Sub MergeVSFHead(ByRef strNew As String, ByVal strConst As String, ByVal strRegsiter As String)
'���ܣ��ϲ�VSF��ͷ�ִ���strConst�У���ΪstrRegsiter��ӣ�strConstû�У�strRegsiter��ɾ��
'������
'  strNew���ϲ�����ִ�
'  strConst�������ִ�
'  strRegister��ע�����ִ�

    Dim i As Integer, j As Integer
    Dim arrConst As Variant, arrReg As Variant
    Dim strI As String, strJ As String
    Dim blnFind As Boolean
    
    strNew = strRegsiter
    
    arrConst = Split(strConst, "|")
    arrReg = Split(strRegsiter, "|")
    For i = LBound(arrConst) To UBound(arrConst)
        If Split(arrConst(i), ",")(0) <> "" Then
            strI = Split(arrConst(i), ",")(0)
        Else
            strI = Split(arrConst(i), ",")(1)
        End If
        
        blnFind = False
        For j = LBound(arrReg) To UBound(arrReg)
            If Split(arrReg(j), ",")(0) <> "" Then
                strJ = Split(arrReg(j), ",")(0)
            Else
                strJ = Split(arrReg(j), ",")(1)
            End If
            If strI = strJ Then
                blnFind = True
                Exit For
            End If
        Next
        
        If blnFind = False Then
            strNew = strNew & "|" & arrConst(i)
        End If
    Next
    
    strRegsiter = strNew
    strNew = ""
    
    arrConst = Split(strConst, "|")
    arrReg = Split(strRegsiter, "|")
    For i = LBound(arrReg) To UBound(arrReg)
        If Split(arrReg(i), ",")(0) <> "" Then
            strI = Split(arrReg(i), ",")(0)
        Else
            strI = Split(arrReg(i), ",")(1)
        End If
        
        blnFind = False
        For j = LBound(arrConst) To UBound(arrConst)
            If Split(arrConst(j), ",")(0) <> "" Then
                strJ = Split(arrConst(j), ",")(0)
            Else
                strJ = Split(arrConst(j), ",")(1)
            End If
            If strI = strJ Then
                blnFind = True
                Exit For
            End If
        Next
        
        If blnFind Then
            strNew = strNew & arrReg(i) & "|"
        End If
    Next
    
    strNew = Left(strNew, Len(strNew) - 1)
End Sub

Public Sub SetVSFHead(ByVal vsfObject As VSFlexGrid, ByVal strHead As String)
'--------------------------------
'���ܣ���ʼ��VSFlexGrid�ؼ����ͷ
'������
'  vsfObject��Ŀ��ؼ���
'  strHead�����ͷ�ĳ�ʼ���ִ�
'
'��ʽ�� "����,,3,1000,s|..."
'   Ԫ��1��Keyֵ��
'   Ԫ��2��Captionֵ��Ĭ��ΪKeyֵ����
'   Ԫ��3�������ԣ�0���ڲ���ʾ�����ƶ���1���ڲ����أ������ƶ���������ʾ��2���û����أ�3���û���ʾ(Ĭ��ֵ)��
'   Ԫ��4���п�ȣ�Ĭ��0����
'   Ԫ��5����ʾ��ʽ��s(Ĭ��)���ַ����� n�����֣� d�����ڣ� t��ʱ�䣻 dt������ʱ��
'--------------------------------
    Dim arrCols As Variant, arrRows As Variant
    Dim i As Integer
    
    On Error GoTo errHandle
    
    arrRows = Split(strHead, "|")
    With vsfObject
        If .Rows = 0 Then .Rows = 1
        .Cols = UBound(arrRows) + 1
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
                    .ColData(i) = Val(arrCols(2))
                End If
                
                '��4Ԫ�أ����
                .ColWidth(i) = Val(arrCols(3))
                
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
                    End If
                Else
                    .ColAlignment(i) = flexAlignLeftCenter
                End If
                
                '������
                If Val(arrCols(2)) = 1 Or Val(arrCols(2)) = 2 Or Val(arrCols(2)) = 0 Then
                    .ColHidden(i) = True
                Else
                    .ColHidden(i) = False
                End If
                
            End If
        Next
        
        If .Cols > 0 Then .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
    Exit Sub
    
errHandle:
    MsgBox Err.Description, vbInformation, gstrSysName
End Sub

Public Function GetCurrentVSFHead(ByVal vsfObject As VSFlexGrid) As String
'-------------------------------------
'���ܣ���ȡVSFĿ��ؼ���ǰ�ı��ͷ�ִ�
'������vsfObject��Ŀ��ؼ�
'���أ����ͷ�ִ�
'-------------------------------------
    Dim i As Integer
    Dim strHead As String, strCol As String
    
    With vsfObject
        strHead = ""
        For i = 0 To .Cols - 1
            '��1Ԫ�أ�Key
            strCol = .ColKey(i) & ","
            '��2Ԫ�أ�Caption
            If strCol = .TextMatrix(0, i) & "," Then
                strCol = strCol & ","
            Else
                strCol = strCol & .TextMatrix(0, i) & ","
            End If
            '��3Ԫ�أ�������
            If Val(.ColData(i)) = 3 Then
                If .ColHidden(i) Then
                    strCol = strCol & "2,"
                Else
                    strCol = strCol & ","
                End If
            Else
                If .ColHidden(i) = False And Val(.ColData(i)) = 2 Then
                    strCol = strCol & "3,"
                Else
                    strCol = strCol & .ColData(i) & ","
                End If
            End If
            '��4Ԫ�أ��п�
            If Val(.ColWidth(i)) = 0 Then
                strCol = strCol & ","
            Else
                strCol = strCol & .ColWidth(i) & ","
            End If
            '��5Ԫ�أ���ʾ��ʽ
            If Trim(.ColFormat(i)) = "" Then
                If .ColAlignment(i) = flexAlignRightCenter Then
                    strCol = strCol & "n"
                Else
                    strCol = Left(strCol, Len(strCol) - 1)
                End If
            Else
                If .ColFormat(i) = "yyyy-mm-dd" Then
                    strCol = strCol & "d"
                ElseIf .ColFormat(i) = "hh:mm:ss" Then
                    strCol = strCol & "t"
                ElseIf .ColFormat(i) = "yyyy-mm-dd hh:mm:ss" Then
                    strCol = strCol & "dt"
                End If
            End If
            '�������
            strHead = strHead & strCol & IIf(i = .Cols - 1, "", "|")
        Next
    End With
    GetCurrentVSFHead = strHead
End Function

Public Sub FillVSFData(ByRef vsfVar As VSFlexGrid, ByRef rsVar As ADODB.Recordset)
'���ܣ�����¼����������������vsf�ؼ���
'������
'  vsfVar��Ҫ������ݵ�Vsf�ؼ�
'  rsVar����¼������

    If rsVar Is Nothing Then Exit Sub
    If rsVar.State <> adStateOpen Then Exit Sub
    If vsfVar Is Nothing Then Exit Sub
    
    Dim i As Integer, intCol As Integer
    Dim lngRow As Long
    
    With rsVar
        vsfVar.Redraw = flexRDNone
        vsfVar.Rows = .RecordCount + 1
        vsfVar.Clear 1
        
        lngRow = 1
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
            For i = 0 To .Fields.Count - 1
                intCol = vsfVar.ColIndex(.Fields(i).Name)
                If intCol >= 0 Then
                    'vsf�д��ڸ��ֶ�
                    vsfVar.TextMatrix(lngRow, intCol) = zlCommFun.NVL(.Fields(i).Value)
                End If
            Next
            
            lngRow = lngRow + 1
            .MoveNext
        Loop
        vsfVar.Redraw = flexRDDirect
    End With

End Sub

Public Sub SetTextMaxLen(ByRef txtVal As TextBox, ByVal strTableField As String)
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = zlStr.FormatString("Select [2] as �ֶ� From [1] Where Rownum < 1 ", _
                        CStr(Split(strTableField, ".")(0)), _
                        CStr(Split(strTableField, ".")(1)))
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ֶ���Ϣ")
    txtVal.MaxLength = rsTmp.Fields(0).DefinedSize
    rsTmp.Close

    Exit Sub
    
errHandle:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Sub

Public Sub SetRecordsetStructure(ByVal bytClass As Byte, ByRef rsVar As ADODB.Recordset)
'���ܣ����ò��ϸ��¼��������ֶ�
'������
'  bytClass��1-���ύ
'  rsVar�����ϸ��¼������
    
    '��ʽ���ֶ���,�ֶ�����,����
    '  3-adInteger��20-adBigInt��200-adVarchar��201-adLongVarchar
    Const STR_NG_PROP     As String = "ҩ��ID;20|�����Ŀ;200;100|ҩƷ����;200;100|��Ʒ��;200;100|���;200;100|��λ;200;100"
    Const STR_SUBMIT_PROP As String = "��ҩҩ��ID;20|�����ĿID;20|����;200;100|���;200;100|ҽ��ID;20|�����;3"
    
    Dim strFieldsProp As String
    Dim arrFields As Variant, arrProp As Variant
    Dim i As Integer

    If Not rsVar Is Nothing Then
        If rsVar.State <> adStateClosed Then rsVar.Close
        Set rsVar = Nothing
    End If
    
    Set rsVar = New ADODB.Recordset
    
    With rsVar
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockPessimistic
        
        If bytClass = 1 Then
            strFieldsProp = STR_SUBMIT_PROP
        End If
        
        '�½�
        arrFields = Split(strFieldsProp, "|")
        For i = LBound(arrFields) To UBound(arrFields)
            arrProp = Split(arrFields(i), ";")
            Select Case Val(arrProp(1))
                Case DataTypeEnum.adVarChar
                    If UBound(arrProp) >= 2 Then
                        .Fields.Append arrProp(0), adVarChar, Val(arrProp(2))
                    Else
                        .Fields.Append arrProp(0), adVarChar, 100           'Ĭ�ϳ���
                    End If
                Case DataTypeEnum.adLongVarChar
                    .Fields.Append arrProp(0), adLongVarChar                'LongVarchar��̬����
                Case DataTypeEnum.adBigInt
                    .Fields.Append arrProp(0), adBigInt
                Case DataTypeEnum.adInteger
                    .Fields.Append arrProp(0), adInteger
            End Select
        Next
        .Open
    End With
End Sub

Public Function SetPublicFontSize(ByRef frmMe As Object, ByVal bytSize As Byte, Optional ByVal strOther As String)
'���ܣ����ô��弰���пؼ��������С
'������frmMe=��Ҫ��������Ĵ������
'      bytSize:����Ϊ9������,0:����Ϊ9������,1,����Ϊ12������
'      strOther:�������������õĿؼ��������ļ���,��ʽΪ����������1,��������2,��������3,....
'˵����1.����漰��VsFlexGrid�ȱ��ؼ�����Ҫ�������ڵĻ������µ����п���и�
'      2.�������δ�г��������ؼ����Զ���ؼ�,��Ҫ���ض�����ָ�������С����ش���ģ������ⵥ������

    Dim objCtrol As Control
    Dim CtlFont As StdFont
    Dim i As Long, lngOldSize As Long
    Dim lngFontSize As Long
    Dim dblRate As Double
    Dim blnDo As Boolean
    Dim strContainer As String
    
    lngFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    frmMe.FontSize = lngFontSize
    strOther = "," & strOther & ","
    blnDo = False
        
    For Each objCtrol In frmMe.Controls
        Select Case TypeName(objCtrol)
            Case "TabStrip", "Label", "ComboBox", "ListView", "OptionButton", "CheckBox", "DTPicker", "TextBox", "SpeedButton", _
                "DockingPane", "CommandBars", "TabControl", "CommandButton", "Frame", "RichTextBox", "MaskEdBox", "IDKindNew", _
                "VSFlexGrid", "StatusBar"
                blnDo = True
            Case Else
                blnDo = False
        End Select
        
        If strOther <> ",," And blnDo Then
            '����CommandBars�û��Զ���ؼ���ȡobjCtrol.Container�����
            strContainer = ""
            On Error Resume Next
            strContainer = objCtrol.Container.Name
            Err.Clear: On Error GoTo 0
            If InStr(1, strOther, "," & strContainer & ",") > 0 Then
                 blnDo = False
            End If
        End If
        
        If blnDo Then
            Select Case TypeName(objCtrol)
                Case "TabStrip"
                        objCtrol.Font.Size = lngFontSize
                Case "Label"
                        If Not LCase(objCtrol.Name) Like "*_fixed" Then
                            lngOldSize = objCtrol.Font.Size
                            dblRate = lngFontSize / lngOldSize
                            
                            objCtrol.Font.Size = lngFontSize
                            objCtrol.Height = frmMe.TextHeight("��") + 20
                            'Label�����Ҫ���е���
                        End If
               Case "ComboBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "ListView"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        For i = 1 To objCtrol.ColumnHeaders.Count
                            objCtrol.ColumnHeaders(i).Width = objCtrol.ColumnHeaders(i).Width * dblRate
                        Next
                Case "OptionButton"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = frmMe.TextWidth("����" & objCtrol.Caption)
                        objCtrol.Height = objCtrol.Height * dblRate
                Case "CheckBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "DTPicker"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = frmMe.TextWidth("2012-01-01    ")
                        objCtrol.Height = frmMe.TextHeight("��") + IIf(bytSize = 0, 100, 120)
                Case "TextBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                        objCtrol.Height = frmMe.TextHeight("��")
                Case "MaskEdBox"
                        objCtrol.FontSize = lngFontSize
                        objCtrol.Width = frmMe.TextWidth(objCtrol.Mask)
                        objCtrol.Height = frmMe.TextHeight("��")
'                Case "ReportControl"
'                        lngOldSize = objCtrol.PaintManager.TextFont.Size
'                        dblRate = lngFontSize / lngOldSize
'
'                        Set CtlFont = objCtrol.PaintManager.CaptionFont
'                        CtlFont.Size = lngFontSize
'                        Set objCtrol.PaintManager.CaptionFont = CtlFont
'                        Set CtlFont = objCtrol.PaintManager.TextFont
'                        CtlFont.Size = lngFontSize
'                        Set objCtrol.PaintManager.TextFont = CtlFont
'                        For Each objrptCol In objCtrol.Columns
'                            objrptCol.Width = objrptCol.Width * dblRate
'                        Next
'                        objCtrol.Redraw
                Case "SpeedButton"
                        Dim objFont As New StdFont
                        
                        Set objFont = frmMe.Font
                        If bytSize = 0 Then
                            objFont.Size = 12
                            dblRate = 0.8
                        Else
                            objFont.Size = 15.75
                            dblRate = 1 / 0.8
                        End If
                        Set objCtrol.Font = objFont
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "VSFlexGrid"
                        Set objCtrol.Font = frmMe.Font
                        objCtrol.Font.Size = IIf(bytSize = 0, 9, 12)
                Case "DockingPane"
                        Set CtlFont = objCtrol.PaintManager.CaptionFont
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.CaptionFont = CtlFont
                        
                        Set CtlFont = objCtrol.TabPaintManager.Font
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.TabPaintManager.Font = CtlFont
        
                        Set CtlFont = objCtrol.PanelPaintManager.Font
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PanelPaintManager.Font = CtlFont
                Case "CommandBars"
                        Set CtlFont = objCtrol.Options.Font
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.Options.Font = CtlFont
                Case "TabControl"
                        Set CtlFont = objCtrol.PaintManager.Font
                        If CtlFont Is Nothing Then  '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.Font = CtlFont
                        objCtrol.PaintManager.Layout = xtpTabLayoutAutoSize
                Case "CommandButton"
                        If Not LCase(objCtrol.Name) Like "*_fixed" Then
                            lngOldSize = objCtrol.FontSize
                            dblRate = lngFontSize / lngOldSize
    
                            objCtrol.FontSize = lngFontSize
                            objCtrol.Width = dblRate * objCtrol.Width
                            objCtrol.Height = dblRate * objCtrol.Height
                        End If
                Case "Frame"
                        objCtrol.FontSize = lngFontSize
                Case "IDKindNew"
                        objCtrol.FontSize = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
                Case "StatusBar"
                        objCtrol.Font.Size = lngFontSize
            End Select
        End If
    Next
End Function

