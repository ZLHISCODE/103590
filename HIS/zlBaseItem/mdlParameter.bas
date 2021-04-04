Attribute VB_Name = "mdlParameter"
Option Explicit
Public Const gstrParSplit1 As String = "^"  '�������������ģ�顢�����š�����ֵ�ָ���
Public Const gstrParSplit2 As String = "#"  '�������������ָ���

Public Enum Enum_Module
    P������Ժ���� = 1131
    p����������� = 1132
    p������Ϣ���� = 1101
    p���ﲡ������ = 1250
    pסԺ�������� = 1251
    p����ҽ���´� = 1252
    pסԺҽ���´� = 1253
    pסԺҽ������ = 1254
    p�����¼���� = 1255
    p�ٴ�·��Ӧ�� = 1256
    p����·��Ӧ�� = 1248
    p�ٴ�·������ = 1078
    p�����¼���� = 1256
    pҽ�����ѹ��� = 1257
    p���Ʊ������ = 1258
    p����ҽ��վ = 1260
    pסԺҽ��վ = 1261
    pסԺ��ʿվ = 1262
    pҽ������վ = 1263
    p������Һ���� = 1264
    p�°�סԺ��ʿվ = 1265
    p������ϲο� = 1270
    pҩƷ���Ʋο� = 1271
    p���˲������� = 1273
    p��Ƭ���߹��� = 1289
    
    'ҩƷҵ��
    pҩƷĿ¼���� = 1023
    pҩƷ�⹺���� = 1300
    pҩƷ������� = 1301
    pҩƷ�ƿ���� = 1304
    pҩƷ���ù��� = 1305
    pҩƷ�̵���� = 1307
    pҩƷ���۹��� = 1333
    pҩƷ�������� = 1331
    pҩƷ������ҩ = 1341
    pҩƷ���ŷ�ҩ = 1342
    pҩƷ������� = 1343
    p�󴦷���� = 1347
    p��Һ�������� = 1345

    '����ҵ��
    p�����ڲ����� = 1070
    p���Ӳ������ = 1560
    p���Ӳ������� = 1561
    p���Ӳ������� = 1562
    
    '����ҵ��
    p��������ģ�� = 9000
    pԤ������� = 1103
    pҽ�ƿ����� = 1107
    p�ҺŰ��� = 1110
    p�ҺŹ��� = 1111
    p������� = 1113
    p�ٴ����ﰲ�� = 1114
    p���ﻮ�۹��� = 1120
    p�����շѹ��� = 1121
    p������ʹ��� = 1122
    p���ﲹ���� = 1124
    pסԺ���ʹ��� = 1133
    p���ҷ�ɢ���� = 1134
    pҽ�����Ҽ��� = 1135
    pסԺ���ʲ��� = 1150
    p���˽��ʹ��� = 1137
    pִ�еǼǹ��� = 1142
    p������˹��� = 1143
    pһ��ͨ���Ѳ��� = 1151
    p�շѲ����� = 1500
    pƱ��ʹ�ü�� = 1501
    p��Ա������ = 1502
    p���ѿ����� = 1503
    pƱ�������� = 1504
    p�շ����ʹ��� = 1506
    
    '�������
    p���ﴦ����� = 1351
    pסԺҩ����� = 1352
    p���������Ŀ = 1353
    p����������� = 1354
    p�������ͳ�� = 1355

    'PACS
    pӰ���Ƭ���� = 1288
    pӰ��ҽ������ = 1290
    pӰ��ɼ����� = 1291
    pӰ�������� = 1294
    p����鵵���� = 1295
    p����軹���� = 1296
End Enum

Public Enum ParaErrType
    PET_���� = 0
    PET_������ʧ = 1 '�ò���������
    PET_�������� = 2 '�ò�����˽�л򱾻������޷��ڴ˴����в�������
    PET_ֵ���� = 3 '�ò���ֵ�����ɿؼ�����Χ
End Enum

Public Sub InitSCBItem(ByRef scb As ShortcutBar, ByVal strItems As String, ByRef lngTPLhwnd As Long, Optional ByVal lngSelectedItem As Long = 1)
'���ܣ���ʼ��һ������������б�
'������
'      strItems         - ��������б����ƣ��Զ��ŷָ�,�����������ݳ�ʼ,���������,�ӿ�����
'      lngTPLhwnd       - �����б��ϰ󶨵�TaskPanel���ڵ���������������Picture��
'      lngSelectedItem  - ȱʡѡ��������,��1��ʼ

    Dim scbItem As ShortcutBarItem
    Dim i As Long
    Dim arrItem As Variant
    
    arrItem = Split(strItems, ",")
    For i = 0 To UBound(arrItem)
        Set scbItem = scb.AddItem(i + 1, arrItem(i), lngTPLhwnd)    'ͼ����ű�ָ����С1������Ҫ��1
        If i + 1 = lngSelectedItem Then Set scb.Selected = scbItem
    Next
    
    scb.ExpandedLinesCount = scb.ItemCount
End Sub


Public Sub InitTPLItem(ByRef scc As ShortcutCaption, ByRef tplFunc As TaskPanel, _
        ByVal strCategory As String, ByVal strItems As String, Optional ByVal lngSelectedItem As Long = 1)
'���ܣ���ʼ�����¼���һ����������б���һ�����飩
'������
'      strCategory      - ��ʾ��ShotcutCaption�ϵĵ�ǰ��������
'      strItems         - ���������������ƣ��Էֺŷָ�,�Զ��ŷָ�ͼ��ID���������鼰������������,����401,1,���ﻮ�۹���;412,2,�����շѹ���;......
'      lngSelectedItem  - ȱʡѡ��������,��1��ʼ

    Dim tplGroup As TaskPanelGroup
    Dim tplItem As TaskPanelGroupItem
    Dim arrItem As Variant
    Dim i As Long
    Dim lngImg As Long, lngId As Long
    Dim strItem As String
    Dim lngUbound As Long
    
    '����һ�����ط���
    scc.Caption = strCategory
    If tplFunc.Groups.Count = 0 Then
        Set tplGroup = tplFunc.Groups.Add(1, "����")
        tplGroup.CaptionVisible = False
        tplGroup.Expanded = True
        
        tplFunc.SetMargins 1, 2, 0, 2, 2
        tplFunc.SetIconSize 24, 24
        tplFunc.SelectItemOnFocus = True
    Else
        Set tplGroup = tplFunc.Groups(1)    'index�Ǵ�1��ʼ��
        tplGroup.Items.Clear
    End If
    
    arrItem = Split(strItems, ";")
    lngUbound = UBound(arrItem)
    For i = 0 To lngUbound
        lngImg = Split(arrItem(i), ",")(0) + 1  'ͼ����ű�ָ����С1������Ҫ��1
        lngId = Split(arrItem(i), ",")(1)       'ID����Ϊ�����ؼ�������Picture�����ţ�
        strItem = Split(arrItem(i), ",")(2)
        Set tplItem = tplGroup.Items.Add(lngId, strItem, xtpTaskItemTypeLink, lngImg)
        If i = lngUbound Then tplItem.SetMargins 0, 0, 0, 0 '��Ȼ���һ��ѡ��ʱ�Ŀ������ȫ��ס����
        If i + 1 = lngSelectedItem Then tplItem.Selected = True: tplFunc.Tag = lngId
    Next
    
End Sub

Public Sub LocatePar(ByRef txtInput As TextBox, ByRef objForm As Form)
'���ܣ����Ҳ�������λ����ʾ
        Dim ctlTmp  As Control, strName As String
        Dim strInput As String, strOldColor As String
        Dim i As Long, p As Long, blnFind As Boolean
        Dim lngStart As Long, lngCount As Long
        Dim objPicPar As PictureBox
        Dim objTarget As Object
         
        lngStart = Val(txtInput.Tag)
        If lngStart = 0 Then lngStart = 1
        strInput = "*" & Trim(txtInput.Text) & "*"
      
        For Each ctlTmp In objForm.Controls
            
            lngCount = lngCount + 1
            If lngCount > lngStart Then
                strName = TypeName(ctlTmp)
                Select Case strName
                Case "Label", "CheckBox", "OptionButton", "Frame"
                    If ctlTmp.Caption Like strInput Then
                        blnFind = True
                        txtInput.Tag = lngCount
                        
                        '��ʱ���֧������
                        If ctlTmp.Container.Name = "picPar" Then
                            Set objPicPar = ctlTmp.Container
                        Else
                            On Error Resume Next    '��������������У��ؼ�����û����ô�༶����
                            If ctlTmp.Container.Container.Name = "picPar" Then
                                Set objPicPar = ctlTmp.Container.Container
                            ElseIf ctlTmp.Container.Container.Container.Name = "picPar" Then
                                Set objPicPar = ctlTmp.Container.Container.Container
                            End If
                            On Error GoTo 0: Err.Clear
                        End If
                        
                        If Not objPicPar Is Nothing Then
                            If objPicPar.Visible = False Then
                                For Each objTarget In objForm.picPar
                                    If objTarget Is objPicPar Then
                                        objTarget.Visible = True
                                        Call objForm.LocateFuncItem(objTarget.Index)
                                    Else
                                        objTarget.Visible = False
                                    End If
                                Next
                                objForm.Refresh
                            End If
                        End If
                        strOldColor = ctlTmp.ForeColor
                        
                        ctlTmp.ForeColor = vbRed
                        ctlTmp.Refresh
                        Call OS.Wait(400)
                        ctlTmp.ForeColor = &H80000012
                        ctlTmp.Refresh
                        Call OS.Wait(200)
                        ctlTmp.ForeColor = vbRed
                        ctlTmp.Refresh
                        Call OS.Wait(400)
                        ctlTmp.ForeColor = &H80000012
                        ctlTmp.Refresh
                        Call OS.Wait(200)
                        ctlTmp.ForeColor = vbRed
                        ctlTmp.Refresh
                        Call OS.Wait(400)
                        
                        ctlTmp.ForeColor = strOldColor
                        ctlTmp.Refresh
                        Exit For
                    End If
                End Select
            End If
        Next
        
        If blnFind = False Then
            If lngStart = 1 Then
                MsgBox "û���ҵ�ƥ��Ĳ�����������������ݡ�", vbInformation, "��������"
            Else
                MsgBox "ȫ�������ˣ�����û���ˡ�", vbInformation, "��������"
                txtInput.Tag = ""
            End If
            
            txtInput.SelStart = 0
            txtInput.SelLength = Len(txtInput.Text)
            If txtInput.Enabled Then txtInput.SetFocus
        End If
End Sub


Public Sub EnterNextCell(ByRef vsobj As VSFlexGrid)
'���ܣ����λ����һ��
    With vsobj
        If .Col + 1 > .Cols - 1 Then
            If .Row + 1 > .Rows - 1 Then .AddItem ""
            .Row = .Row + 1: .Col = .FixedCols
        Else
            .Col = .Col + 1
        End If
        '�������������ݹ��ٶ�λ����һ��λ��
        If .ColHidden(.Col) = True Then Call EnterNextCell(vsobj)
        .ShowCell .Row, .Col
    End With
End Sub

Public Function GetPar(ByRef rsPar As ADODB.Recordset, Optional ByVal strModules As String) As ADODB.Recordset
'���ܣ���ȡϵͳ������ָ����ģ����������ؼ�¼��
'������ģ��Ŵ�
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    strSQL = "Select ID,������,������,Nvl(����ֵ,ȱʡֵ) as ����ֵ,NVL(����, 0) As ����,Nvl(˽��, 0) ˽��, NVL(����, 0) ����,Ӱ�����˵��,����˵��,����˵��,����˵��,Decode(����˵��,Null,0,1) as �Ƿ�ؼ�����,Nvl(ģ��,0) as ģ�� " & vbCrLf & _
            "From Zlparameters Where ϵͳ = " & glngSys & "  And Nvl(����,0) = 0 And " & _
            IIF(strModules = "", "ģ�� Is Null", "(ģ�� Is Null Or ģ�� In(" & strModules & "))")
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "��ȡϵͳ����")
    Set rsPar = zlDatabase.CopyNewRec(rsTmp, False, "", Array("������ֵ", adVarChar, 4000, Empty, "�޸�״̬", adInteger, 1, Empty, _
                "�ؼ�����", adVarChar, 50, Empty, "�ؼ��������", adInteger, 3, Empty, "�ؼ���ʶ", adVarChar, 50, Empty, "ErrType", adInteger, 1, Empty))
    '���˽�б�������
    Call rec.Update(rsPar, "(˽��=1 And ����=0) OR (����=1 And ����=0)", "ErrType", PET_��������)
    Set GetPar = rsTmp
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub SetParToControl(ByVal strPar As String, ByRef rsPar As ADODB.Recordset, ByRef arrObj As Variant, Optional ByVal bytMode As Byte = 0)
'���ܣ����ò���ֵ�����ü���ؼ�,������rsPar�н���������ؼ����Ƽ�������ŵĹ���
'������strPar    -ģ���1:������1(�������1):�ؼ����1,ģ���2:������2(�������2):�ؼ����2,......����arrObjΪ��������ʱ����
'      rsPar    -�����ݿ��ȡ�Ĳ�����¼��
'      arrObj   -֧��Checkbox,ComboBox,UpDown,OptionButton,ListBox,TextBox�ؼ�����(Ҫ�����������)
'               -����Ƕ������飬���ʽΪ��ģ��1,������1,�ؼ�����1,ģ��2,������2,�ؼ�����2,......
'      bytMode- ListBox��ItemDataȡֵģʽ��0-��Chrת���޷ָ���1-ֱ���ö��ŷָ�(*��ʾȫ��ƥ��),2-List(�ı�),3-ƥ��Ĳ���ѡ,4-���ŷָ�(����ȫѡ)
'               ComboBox��ȡֵģʽ��0-ȡListIndex,1-ȡItemData,2-val(List(i)),3-List(i)�ı��Ƚ�
'               OPtionButton 0-����������Index��1-��������һλ��ΪIndex
    Dim strMsg As String, strErr As String
    Dim arrPar As Variant, i As Long, j As Long
    Dim lngModule As Long, lngPar As Long, strParName As String, lngObjIndex As Long
    Dim strType As String, strCtrlName As String
    Dim objTmp As Object
    
    On Error Resume Next
    
    If IsArray(arrObj) Then     'OptionButton����ؼ�������
        For i = 0 To UBound(arrObj) Step 3
            lngModule = arrObj(i)
            strParName = arrObj(i + 1)
            If IsNumeric(strParName) Then
                lngPar = Val(strParName): strParName = ""
            Else
                lngPar = 0
            End If
            Set objTmp = arrObj(i + 2)
            
            rsPar.Filter = IIF(strParName <> "", "������='" & strParName & "'", "������=" & lngPar) & " And ģ�� = " & lngModule
            strType = TypeName(objTmp(0))
            strCtrlName = objTmp(0).Name
            If Err.Number <> 0 Then Err.Clear
            If rsPar.RecordCount > 0 Then
                If strType = "OptionButton" Then
                    rsPar!�ؼ����� = strCtrlName
                    If bytMode = 0 Then
                        objTmp(Val("" & rsPar!����ֵ)).value = True
                    ElseIf bytMode = 1 Then
                        If "" & rsPar!����ֵ <> "" Then
                            objTmp(Val(Mid("" & rsPar!����ֵ, 1, 1))).value = True
                        Else
                            objTmp(0).value = True
                        End If
                    End If
                    If Err.Number <> 0 Then
                        Err.Clear
                        If Val(rsPar!ErrType & "") = 0 Then rsPar!ErrType = PET_ֵ���� 'ֵ��Χ����ȷ
                    End If
                End If
                rsPar!�ؼ�������� = 0
                rsPar.Update
            Else
                '���Ӷ�ʧ�Ĳ���
                rsPar.AddNew Array("ID", "ģ��", "������", "������", "�ؼ�����", "�ؼ��������", "ErrType"), Array(-1, lngModule, lngPar, strParName, IIF(strType = "OptionButton", strCtrlName, Null), 0, PET_������ʧ)
            End If
        Next
    Else
        arrPar = Split(strPar, ",")
        strType = TypeName(arrObj(0))
        
        For i = 0 To UBound(arrPar)
            lngModule = Split(arrPar(i), ":")(0)
            strParName = Split(arrPar(i), ":")(1)
            If IsNumeric(strParName) Then
                lngPar = Val(strParName): strParName = ""
            Else
                lngPar = 0
            End If
            lngObjIndex = Split(arrPar(i), ":")(2)
            strCtrlName = arrObj(lngObjIndex).Name
            If strType = "UpDown" Then strCtrlName = "txtUD"
            rsPar.Filter = IIF(strParName <> "", "������='" & strParName & "'", "������=" & lngPar) & " And ģ�� = " & lngModule
            If rsPar.RecordCount > 0 Then
                rsPar!�ؼ����� = strCtrlName
                Select Case strType
                    Case "CheckBox"
                        arrObj(lngObjIndex).value = IIF(Val("" & rsPar!����ֵ) <> 0, 1, 0)
                    Case "ComboBox"
                        If bytMode = 0 Then
                            arrObj(lngObjIndex).ListIndex = Val("" & rsPar!����ֵ)
                        Else
                            With arrObj(lngObjIndex)
                                For j = 0 To .ListCount - 1
                                    If bytMode = 1 Then
                                        If .ItemData(j) = Val("" & rsPar!����ֵ) Then
                                            .ListIndex = j
                                            Exit For
                                        End If
                                    ElseIf bytMode = 3 Then '�ı��Ƚ�
                                        If .List(j) = NVL(rsPar!����ֵ) Then
                                            .ListIndex = j: Exit For
                                        End If
                                    Else
                                        If Val(.List(j)) = Val("" & rsPar!����ֵ) Then
                                            .ListIndex = j
                                            Exit For
                                        End If
                                    End If
                                Next
                                If .ListCount > 0 And j > .ListCount - 1 Then .ListIndex = 0
                            End With
                        End If
                        arrObj(lngObjIndex).Tag = bytMode
                    Case "UpDown"
                        arrObj(lngObjIndex).value = rsPar!����ֵ
                    Case "OptionButton"
                        arrObj(Val(rsPar!����ֵ)).value = True
                        lngObjIndex = 0  '����Ź̶��洢Ϊ0
                    Case "TextBox"
                        arrObj(lngObjIndex).Text = NVL(rsPar!����ֵ)
                    Case "ListBox"
                        For j = 0 To arrObj(lngObjIndex).ListCount - 1
                            If bytMode = 0 Then
                                If InStr("" & rsPar!����ֵ, Chr(arrObj(lngObjIndex).ItemData(j))) > 0 Then
                                    arrObj(lngObjIndex).Selected(j) = True
                                End If
                            ElseIf bytMode = 1 Then
                                If "" & rsPar!����ֵ = "*" Or InStr("," & rsPar!����ֵ & ",", "," & arrObj(lngObjIndex).ItemData(j) & ",") > 0 Then
                                    arrObj(lngObjIndex).Selected(j) = True
                                End If
                            ElseIf bytMode = 3 Then
                                If InStr("" & rsPar!����ֵ, arrObj(lngObjIndex).ItemData(j)) = 0 Then
                                    arrObj(lngObjIndex).Selected(j) = True
                                End If
                            ElseIf bytMode = 4 Then
                                If InStr("," & rsPar!����ֵ & ",", "," & arrObj(lngObjIndex).ItemData(j) & ",") > 0 Then
                                    arrObj(lngObjIndex).Selected(j) = True
                                End If
                            Else
                                If InStr("," & rsPar!����ֵ & ",", "," & arrObj(lngObjIndex).List(j) & ",") > 0 Then
                                    arrObj(lngObjIndex).Selected(j) = True
                                End If
                            End If
                        Next
                        arrObj(lngObjIndex).Tag = bytMode
                    Case Else
                        If Val(rsPar!ErrType & "") = 0 Then rsPar!ErrType = PET_ֵ���� 'ֵ��Χ����ȷ
                End Select
                
                rsPar!�ؼ�������� = lngObjIndex
                rsPar.Update
                If Err.Number <> 0 Then
                    Err.Clear
                    If Val(rsPar!ErrType & "") = 0 Then rsPar!ErrType = PET_ֵ���� 'ֵ��Χ����ȷ
                End If
            Else
                '���Ӷ�ʧ�Ĳ���
                rsPar.AddNew Array("ID", "ģ��", "������", "������", "�ؼ�����", "�ؼ��������", "ErrType"), Array(-1, lngModule, lngPar, strParName, strCtrlName, IIF(strType = "OptionButton", 0, lngObjIndex), PET_������ʧ)
            End If
        Next
    End If
'    rsPar.Filter = ""
End Sub

Public Sub SetParRelation(ByRef arrObj As Variant, ByVal lngObjIndex As Long, ByRef rsPar As ADODB.Recordset, _
                        Optional ByVal varPar As Variant, Optional ByVal lngModule As Long, _
                        Optional ByVal strObjTag As String, Optional strGridCol As String = "", _
                        Optional ByVal blnNotClearIndex As Boolean = False)
'���ܣ����ò�����ؼ��Ĺ������Ա�ؼ�������ʾʱ���ݵ�ǰ�ؼ������Ҳ�������ʾ˵����Ϣ���Լ����ڹؼ������ľ�����ʾ
'������varPar   -�����Ź�����������ֵΪ0���ʱ�����µ�ǰλ�ü�¼
'      lngModule-ģ��ţ���ֵΪ0ʱ����ʾϵͳ����
'      strGridCol-�󶨱������
'      blnNotClearIndex-���������ֵ��������:lngObjIndex(�ؼ��������)��
    Dim strType As String, strObjName As String
    Dim lngPar As Long, strParName As String
    If TypeName(varPar) <> "Error" Then
        strParName = varPar & ""
        If IsNumeric(strParName) Then
            lngPar = Val(varPar)
            strParName = ""
        End If
    End If
    If lngPar <> 0 Or strParName <> "" Then
        rsPar.Filter = IIF(strParName <> "", "������='" & strParName & "'", "������=" & lngPar) & " And ģ�� = " & lngModule
    End If
    
    strType = TypeName(arrObj)
    If strType = "Object" Then  '�ؼ�����
        strObjName = arrObj(lngObjIndex).Name
        strType = TypeName(arrObj(lngObjIndex))
        If strType = "OptionButton" Then lngObjIndex = 0    '�����ʱ��̶���0������ǿ��ָ��Ϊ0�����ܴ���ֵ
    Else
        strObjName = arrObj.Name
        If blnNotClearIndex = False Then lngObjIndex = 0
    End If
    
    rsPar!�ؼ����� = strObjName & strGridCol
    rsPar!�ؼ�������� = lngObjIndex
    rsPar!�ؼ���ʶ = strObjTag
    rsPar.Update
End Sub


Public Sub SetParChange(ByRef arrObj As Variant, ByVal lngObjIndex As Long, ByRef rsPar As ADODB.Recordset, _
                        Optional ByVal blnValue As Boolean, _
                        Optional ByVal strValue As String, Optional strGridCols As String)
'���ܣ������仯ʱ��������ֵ���޸�״̬
'������blnValue-ָ���οؼ���Ӧ�Ĳ���ֵ
'      strValue-�������ϲ��������⴦��Ĳ���,�������ֵ���޷�ֱ��ͨ���ؼ�ȡֵ��
'      strGridCols-�󶨵���(����ö��ŷ���)
    Dim blnDo As Boolean
    Dim i As Long, strType As String
    Dim str��� As String, blnȫѡ As Boolean
    Dim objTmp As Variant, varTemp As Variant, intCol As Integer
    Dim bytMode As Byte
'       ListBox��ItemDataȡֵģʽ��0-��Chrת���޷ָ���1-ֱ���ö��ŷָ�(*��ʾȫ��ƥ��),2-List(�ı�)
'       ComboBox��ȡֵģʽ��0-ȡListIndex,1-ȡItemData,2-val(List),3-List
    
    strType = TypeName(arrObj)
    
    If strType = "Object" Then  '�ؼ�����
        strType = TypeName(arrObj(lngObjIndex))
        Set objTmp = arrObj(lngObjIndex)
    Else
        Set objTmp = arrObj
    End If
    If strType = "OptionButton" Then lngObjIndex = 0    '�����ʱ��̶���0������ǿ��ָ��Ϊ0�����ܴ���ֵ
    
    rsPar.Filter = "�ؼ����� = '" & objTmp.Name & strGridCols & "' And �ؼ��������=" & lngObjIndex
    If rsPar.RecordCount > 0 Then
        blnDo = True
        If blnValue Then
            rsPar!������ֵ = strValue
        Else
            Select Case strType
                Case "CheckBox"
                    rsPar!������ֵ = objTmp.value
                Case "ComboBox"
                    bytMode = Val(objTmp.Tag)
                    If bytMode = 0 Then
                        rsPar!������ֵ = objTmp.ListIndex
                    ElseIf bytMode = 1 Then
                        rsPar!������ֵ = objTmp.ItemData(objTmp.ListIndex)
                    ElseIf bytMode = 2 Then
                        rsPar!������ֵ = Val(objTmp.List(objTmp.ListIndex))
                    Else
                        rsPar!������ֵ = objTmp.List(objTmp.ListIndex)
                    End If
                Case "TextBox"  'UpDown������txtUD
                    rsPar!������ֵ = objTmp.Text
                Case "OptionButton"
                    For i = 0 To arrObj.UBound
                        If arrObj(i).value Then Exit For
                    Next
                    rsPar!������ֵ = i
                Case "ListBox"
                    blnȫѡ = True
                    bytMode = Val(objTmp.Tag)
                    For i = 0 To objTmp.ListCount - 1
                        If objTmp.Selected(i) Then
                            If bytMode = 0 Then
                                str��� = str��� & Chr(objTmp.ItemData(i))
                            ElseIf bytMode = 1 Then
                                str��� = str��� & "," & objTmp.ItemData(i)
                            ElseIf bytMode = 3 Then
                                '�෴
                            ElseIf bytMode = 4 Then
                                str��� = str��� & "," & objTmp.ItemData(i)
                            Else
                                str��� = str��� & "," & objTmp.List(i)
                            End If
                        Else
                            If bytMode = 3 Then str��� = str��� & "," & objTmp.ItemData(i)
                            blnȫѡ = False
                        End If
                    Next
                    If bytMode = 1 Then
                        str��� = IIF(blnȫѡ, "*", Mid(str���, 2))
                    ElseIf bytMode = 2 Then
                        str��� = Mid(str���, 2)
                    ElseIf bytMode = 3 Then
                        str��� = Mid(str���, 2)
                    ElseIf bytMode = 4 Then
                        str��� = Mid(str���, 2)
                    End If
                    
                    rsPar!������ֵ = str���
                Case Else
                    blnDo = False
            End Select
        End If
        If blnDo Then
            rsPar.Update
            
            If "" & rsPar!������ֵ <> "" & rsPar!����ֵ Then
                rsPar!�޸�״̬ = 1
                If rsPar!�Ƿ�ؼ����� = 1 Then Call MsgBox("���ѣ�" & rsPar!����˵��, vbExclamation, "����")
            Else
                rsPar!�޸�״̬ = 0
            End If
            rsPar.Update
            
            Select Case strType
                Case "CheckBox", "ComboBox", "TextBox", "ListBox", "ListView"
                    objTmp.ForeColor = IIF(Val("" & rsPar!�޸�״̬) = 1, &HC0&, &H0&)             '�޸ĺ������ɫǰ��ɫ��ʶ
                Case "VSFlexGrid"
                    If strGridCols <> "" Then
                        varTemp = Split(strGridCols, ",")
                        For i = 0 To UBound(varTemp)
                            intCol = Val(varTemp(i))
                            objTmp.Cell(flexcpForeColor, objTmp.FixedRows, intCol, objTmp.Rows - 1, intCol) = IIF(Val("" & rsPar!�޸�״̬) = 1, &HC0&, &H0&)
                        Next
                    Else
                        objTmp.ForeColor = IIF(Val("" & rsPar!�޸�״̬) = 1, &HC0&, &H0&)             '�޸ĺ������ɫǰ��ɫ��ʶ
                    End If
                Case "OptionButton"
                    For i = arrObj.LBound To arrObj.UBound
                        On Error Resume Next
                        If i = objTmp.Index Then
                            arrObj(i).ForeColor = IIF(Val("" & rsPar!�޸�״̬) = 1, &HC0&, &H0&)
                        Else
                            arrObj(i).ForeColor = &H0& '�����Ļָ���ɫ
                        End If
                        If Err.Number <> 0 Then Err.Clear
                        On Error GoTo 0
                    Next
            End Select
        End If
    End If
End Sub

Public Sub ShowErrParasMsg(ByRef objFrmMe As Object, ByRef rsPar As ADODB.Recordset)
'���ܣ����������ʾ����������������н��õȱ�ʶ
'������objFrmMe=����
'         rsPar=������¼��
'˵�����ú����ڲ���������ɺ���Ե��ã�ֻ�����һ�μ���
    Dim arrObject As Variant, arrTmp As Variant
    Dim objTmp As Object
    Dim strType As String, strCtrlName As String, blnArray As Boolean
    Dim strMsg As String, strTmp As String, petCurType As ParaErrType
    Dim intCount As Integer
    
    On Error GoTo errH
    '���ý�����ɫ
    rsPar.Filter = "ErrType<>Null And ErrType<>" & PET_ֵ����
    Do While Not rsPar.EOF
        strCtrlName = rsPar!�ؼ����� & ""
        blnArray = False
        On Error Resume Next
        If strCtrlName <> "" Then
            Set arrObject = Nothing: Set objTmp = Nothing
            Set arrObject = CallByName(objFrmMe, strCtrlName, VbGet)
            If Err.Number <> 0 Then Err.Clear
            strType = TypeName(arrObject)
            If TypeName(arrObject) = "Object" Then
                blnArray = True
                For Each objTmp In arrObject
                    strType = TypeName(objTmp)
                    Exit For
                Next
            End If
            If strType <> "Empty" And strType <> "Nothing" Then
                If blnArray Then
                    Set objTmp = arrObject(Val(rsPar!�ؼ�������� & ""))
                Else
                    Set objTmp = arrObject
                End If
                Select Case strType
                    Case "OptionButton"
                        For Each objTmp In arrObject
                            objTmp.ForeColor = &H808080
                            objTmp.Enabled = False
                        Next
                    Case "TextBox", "ComboBox"
                        objTmp.ForeColor = &H808080
                        objTmp.Locked = True
                        If strCtrlName = "txtUD" Then 'ud(UpDown)�ؼ�����
                            Set arrTmp = CallByName(objFrmMe, "ud", VbGet)
                            If Err.Number <> 0 Then Err.Clear
                            Set objTmp = arrTmp(Val(rsPar!�ؼ�������� & ""))
                            objTmp.ForeColor = &H808080
                            objTmp.Enabled = False
                        End If
                    Case Else
                        objTmp.ForeColor = &H808080
                        objTmp.Enabled = False
                End Select
            End If
        End If
        rsPar.MoveNext
    Loop
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo errH
    '��Ϣ��ʾ��
    '����ʾ��ؼ��޹����ı���˽�в���
    rsPar.Filter = "(ErrType=" & PET_�������� & " And �ؼ�����<>Null ) OR (ErrType<>Null And ErrType<>" & PET_�������� & ")"
    rsPar.Sort = "ErrType,ģ��,������,������"
    petCurType = PET_����
    strMsg = "": strTmp = ""
    Do While Not rsPar.EOF
        If petCurType <> Val(rsPar!ErrType) Then
            petCurType = Val(rsPar!ErrType): intCount = 0
            strMsg = strMsg & IIF(strMsg = "", "", vbNewLine) & strTmp
            strTmp = Decode(petCurType, PET_������ʧ, "���²���δ��������ȡ��������ȱ����Щ�������ݣ����鴦��", _
                                                            PET_��������, "���²����������ͱ��Ϊ������˽�в����������ڴ˴����ã��뵽���������ã�", _
                                                            PET_ֵ����, "���²�����ֵ��������ֵ��Χ�����鴦��", "")
        End If
        strTmp = strTmp & IIF(intCount Mod 2 = 0, vbNewLine, ",  ") & IIF(rsPar!ģ�� = 0, "ϵͳ����:", rsPar!ģ�� & "ģ�����:") & IIF(rsPar!������ & "" <> "", rsPar!������, rsPar!������)
        intCount = intCount + 1
        rsPar.MoveNext
    Loop
    If strTmp <> "" Then
        strMsg = strMsg & IIF(strMsg = "", "", vbNewLine) & strTmp
    End If
    If strMsg <> "" Then
        MsgBox strMsg, vbExclamation, "ע��"
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
End Sub

Public Function CheckParChanged(ByRef arrObj As Variant, ByVal lngObjIndex As Long, ByRef rsPar As ADODB.Recordset) As Boolean
'���ܣ�����ָ���Ŀؼ�������ţ��ж϶�Ӧ�Ĳ���ֵ�Ƿ�ı�
    Dim strType As String
    
    strType = TypeName(arrObj)
    If strType = "Object" Then  '�ؼ�����
        strType = TypeName(arrObj(lngObjIndex))
        If strType = "OptionButton" Then lngObjIndex = 0    '�����ʱ��̶���0������ǿ��ָ��Ϊ0�����ܴ���ֵ
        
        rsPar.Filter = "�ؼ�����='" & arrObj(lngObjIndex).Name & "' And �ؼ��������=" & lngObjIndex
    Else
        rsPar.Filter = "�ؼ�����='" & arrObj.Name & "' And �ؼ��������=0"
    End If
    
    If rsPar.RecordCount > 0 Then
        CheckParChanged = (Val("" & rsPar!�޸�״̬) = 1)
    End If
End Function


Public Function GetParOriginalValue(ByRef arrObj As Variant, ByVal lngObjIndex As Long, ByRef rsPar As ADODB.Recordset) As String
'���ܣ�����ָ���Ŀؼ�������ţ����ز���ԭʼֵ
    Dim strType As String
    
    strType = TypeName(arrObj)
    If strType = "Object" Then  '�ؼ�����
        strType = TypeName(arrObj(lngObjIndex))
        If strType = "OptionButton" Then lngObjIndex = 0    '�����ʱ��̶���0������ǿ��ָ��Ϊ0�����ܴ���ֵ
            
        rsPar.Filter = "�ؼ�����='" & arrObj(lngObjIndex).Name & "' And �ؼ��������=" & lngObjIndex
    Else
        rsPar.Filter = "�ؼ�����='" & arrObj.Name & "' And �ؼ��������=0"
    End If
    If rsPar.RecordCount > 0 Then
        GetParOriginalValue = rsPar!����ֵ
    End If
End Function

Public Function SavePar(ByRef rsPar As ADODB.Recordset, ByRef frmParent As Form) As Boolean
'���ܣ������޸Ĺ��Ĳ���
    Dim strPars As String, strPar As String
    
    With rsPar
        'ֻ����û�д���Ĳ�����ֵ������Χ�Ĳ�������
        .Filter = "(�޸�״̬=1 ANd ErrType =Null) OR  (�޸�״̬=1 And ErrType=" & PET_ֵ���� & ")"
        Do While Not .EOF
            If InStr(!������ֵ, gstrParSplit1) > 0 Or InStr(!������ֵ, gstrParSplit2) > 0 Then
                MsgBox "ģ��" & !ģ�� & "�Ĳ���[" & !������ & "]���зǷ���" & gstrParSplit1 & "��" & gstrParSplit2 & "����������!" & vbCrLf & _
                    "����ֵ:" & !������ֵ, vbExclamation, "����"
                Exit Function
            End If
            '����Ŀǰ����ֵ�а�����:,|���ַ�������ѡȡ^#Ϊ�ָ���
            strPar = !ģ�� & gstrParSplit1 & !������ & gstrParSplit1 & !������ֵ
            
            '�ж���ؼ���Ӧһ���������������ȥ���ظ�
            If InStr(strPars, strPar) = 0 Then strPars = strPars & gstrParSplit2 & strPar
            .MoveNext
        Loop
    End With
    strPars = Mid(strPars, 2)
    
    If strPars <> "" Then
        On Error GoTo ErrHandle
        gstrSQL = "zl_Parameters_Update_Batch(" & glngSys & ",'" & strPars & "','" & gstrUserName & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "��������")
        
        '����޸��˹ؼ��������򵯳���������ԭ�򲢱���
        rsPar.Filter = "(�Ƿ�ؼ�����=1 And �޸�״̬=1 ANd ErrType =Null) OR  (�Ƿ�ؼ�����=1 And �޸�״̬=1 And ErrType=" & PET_ֵ���� & ")"
        If rsPar.RecordCount > 0 Then
            Call frmParReason.ShowMe(frmParent, rsPar)
        End If
    End If
    SavePar = True
    Call zlDatabase.ClearParaCache '��ղ�������
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Call zlDatabase.ClearParaCache '��ղ�������
End Function

Public Function GetFuncID(ByVal strName As String, ByRef arrFunc As Variant) As Long
'���ܣ��������Ʒ��ض�������ID
'������arrFunc-������������������飬�Էֺŷָ�,�Զ��ŷָ�ͼ��ID���������鼰������������,����401,1,���ﻮ�۹���;412,2,�����շѹ���;......
'���أ�����ID
    Dim i As Long, j As Long
    Dim arrTmp As Variant
    
    For i = 0 To UBound(arrFunc)
        arrTmp = Split(arrFunc(i), ";")
        For j = 0 To UBound(arrTmp)
            If strName = Split(arrTmp(j), ",")(2) Then
                GetFuncID = Split(arrTmp(j), ",")(1)
                Exit Function
            End If
        Next
    Next
End Function

Public Sub SetParTip(ByRef arrObj As Variant, ByVal lngObjIndex As Long, ByRef rsPar As ADODB.Recordset, _
    Optional ByVal strObjTag As String, Optional ByVal objOtherControl As Object, _
    Optional ByVal strGridCol As String)
'���ܣ����ݿؼ�������ŷ�����֯�õĲ�����ʾ�ı�
'objOtherControl��ָ���������ؼ�����ʾ��ʾ�ı�
'strGridCol-�󶨵���
    Dim strTip As String
    Dim strType As String
    Dim blnArray As Boolean
    Dim petCur As ParaErrType
    strType = TypeName(arrObj)
    If strType = "Object" Then  '�ؼ�����
        blnArray = True
        strType = TypeName(arrObj(lngObjIndex))
        If strType = "OptionButton" Then  '�����ʱ��̶���0������ǿ��ָ��Ϊ0�����ܴ���ֵ
            rsPar.Filter = "�ؼ�����='" & arrObj(lngObjIndex).Name & strGridCol & "' And �ؼ��������=0"
        Else
            rsPar.Filter = "�ؼ�����='" & arrObj(lngObjIndex).Name & strGridCol & "' And �ؼ��������=" & lngObjIndex & IIF(strObjTag <> "", " And �ؼ���ʶ='" & strObjTag & "'", "")
        End If
    Else
        rsPar.Filter = "�ؼ�����='" & arrObj.Name & strGridCol & "' And �ؼ��������=0" & " And �ؼ���ʶ='" & strObjTag & "'"
    End If
    
    If rsPar.RecordCount > 0 Then
        petCur = Val(rsPar!ErrType & "")
        strTip = IIF(rsPar!ģ�� = 0, "ϵͳȫ�ֲ���", "ģ��ţ�" & rsPar!ģ��) & "�������ţ�" & rsPar!������
        
        If petCur = PET_�������� Or petCur = PET_������ʧ Then
            strTip = strTip & vbCrLf & "�������þ���|" & IIF(petCur = PET_��������, "�ò�����ǰ����Ϊ������˽�в����������ڴ˴����ã��뵽���������á�", "�ò���δ��������ȡ��������ȱ�ٸò������ݡ�")
        End If
        strTip = strTip & vbCrLf & "Ӱ�����˵��|" & rsPar!Ӱ�����˵��
        If Not IsNull(rsPar!����˵��) Then strTip = strTip & vbCrLf & "����˵��|" & rsPar!����˵��
        If Not IsNull(rsPar!����˵��) Then strTip = strTip & vbCrLf & "����˵��|" & rsPar!����˵��
        If Not IsNull(rsPar!����˵��) Then strTip = strTip & vbCrLf & "����˵��|" & rsPar!����˵��
    End If
    
    If strTip <> "" Then
        If Not objOtherControl Is Nothing Then
            Call zlCommFun.ShowTipInfo(objOtherControl.hwnd, strTip, True, True, 8800)
        ElseIf blnArray Then
            Call zlCommFun.ShowTipInfo(arrObj(lngObjIndex).hwnd, strTip, True, True, 8800)
        Else
            Call zlCommFun.ShowTipInfo(arrObj.hwnd, strTip, True, True, 8800)
        End If
    End If
End Sub

Public Sub SetPrompt(ByRef lblPrompt As Label, ByVal strPrompt As String)
'���ܣ�������ʾ��Ϣ���Ժ��Զ���ʧ
    lblPrompt.Caption = strPrompt
    lblPrompt.Refresh
    Call OS.Wait(2500)
    lblPrompt.Caption = ""
End Sub


Public Sub SetVsfEditable(ByRef vsf As VSFlexGrid, ByVal blnEdit As Boolean)
'���ܣ����ñ��ؼ��Ŀ����Լ����
    With vsf
        .Enabled = blnEdit
        .Editable = IIF(blnEdit, flexEDKbdMouse, flexEDNone)
        .ForeColor = IIF(blnEdit, vsf.Container.ForeColor, &H808080)
        .BackColor = IIF(blnEdit, &H80000005, vsf.Container.BackColor)
    End With
End Sub

Public Sub SetLstSelected(ByRef lst As ListBox, ByVal blnSel As Boolean)
'���ܣ�ȫѡ��ȫ��ListBox��Ŀ������λ�ò���
    Dim i As Long, Y As Long
    
    With lst
        Y = .ListIndex
        For i = 0 To .ListCount - 1
            .Selected(i) = blnSel    '������lst_ItemCheck�¼�
        Next
        .ListIndex = Y
    End With
End Sub




