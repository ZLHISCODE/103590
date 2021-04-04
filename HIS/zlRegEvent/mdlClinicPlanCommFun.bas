Attribute VB_Name = "mdlClinicPlanCommFun"
Option Explicit
Public Const G_AlternateColor As Long = 16772055   '�н���ɫ
Public Const G_LostFocusColor As Long = &HE0E0E0   'ʧȥ����ʱ�����񱳾�ɫ

Public Enum G_Enum_Fun
    Fun_View = 0 '�鿴
    Fun_Add = 1 '����
    Fun_Update = 2 '�༭
    Fun_Delete = 3 'ɾ��
    Fun_TempPlan = 4 '��ʱ����(�̶������)
    Fun_UpdateUnit = 5 '����������λԤԼ�Һ�
    Fun_AddSignalSourcePlan = 6 '������Դ
    Fun_TempPlanRecord = 7 '��������ʱ����
    Fun_TempPlanVerify = 8 '��ʱ����(�̶������)���
    Fun_TempPlanCancel = 9 '��ʱ����(�̶������)ȡ�����
    Fun_UpdatePlan = 10 '�����ѷ�����İ���
End Enum

Public Enum gRegistPlanEditMode
    ED_RegistPlan_Edit = 0
    ED_RegistPlan_View = 1
    ED_RegistPlan_UpdateUnit = 2
    ED_RegistPlan_NumLimitModify = 3
End Enum

Public Enum PictureTextAlignmentSettings
    'ͼƬ���ı���ʾλ��
    pictxtAlignLeftTop = 0
    pictxtAlignLeftCenter = 1
    pictxtAlignLeftBottom = 2
    pictxtAlignCenterTop = 3
    pictxtAlignCenterCenter = 4
    pictxtAlignCenterBottom = 5
    pictxtAlignRightTop = 6
    pictxtAlignRightCenter = 7
    pictxtAlignRightBottom = 8
End Enum

Public Function GetPlanKey(ByVal strItem As String) As String
    '������Ŀ�������ͻ�ȡKeyֵ
    GetPlanKey = "K" & IIf(IsDate(strItem), Format(strItem, "yyyy-mm-dd"), strItem)
End Function

Public Function HavePrivs(ByVal strPrivs As String, ByVal strMyPriv As String, Optional ByVal blnAnd As Boolean) As Boolean
    '�ж�Ȩ��
    '   strMyPriv ����÷ֺ�";"�ָ�
    '   blnAnd True-���Ȩ����And��ϵ��False-���Ȩ����Or�Ĺ�ϵ
    Dim varPrivs As Variant, i As Integer
    Dim blnHave As Boolean
    
    If InStr(strMyPriv, ";") > 0 Then
        varPrivs = Split(strMyPriv, ";")
        blnHave = IIf(blnAnd, True, False)
        For i = 0 To UBound(varPrivs)
            If blnAnd Then
                If zlStr.IsHavePrivs(strPrivs, varPrivs(i)) = False Then
                    blnHave = False: Exit For
                End If
            Else
                If zlStr.IsHavePrivs(strPrivs, varPrivs(i)) Then
                    blnHave = True: Exit For
                End If
            End If
        Next
    Else
        blnHave = zlStr.IsHavePrivs(strPrivs, strMyPriv)
    End If
    HavePrivs = blnHave
End Function

Public Function ZDate(ByVal varValue As Variant, Optional ByVal varDefault As Variant = "", _
    Optional ByVal blnDataBase As Boolean = True) As String
'���ܣ���ȱʡʱ��ת��Ϊ"NULL"��,������SQL���ʱ��
    Dim strTemp As String
    
    If blnDataBase Then
        If IsDate(varValue) Then
            If DateDiff("s", varValue, "1899-12-30") = 0 Then
                'varValueΪ��
                If IsDate(varDefault) Then
                    If DateDiff("s", varDefault, "1899-12-30") = 0 Then
                        'varDefaultΪ��
                        ZDate = "NULL"
                    Else
                        ZDate = "To_Date('" & Format(varDefault, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                    End If
                Else
                    ZDate = IIf(CStr(varDefault) = "", "NULL", CStr(varDefault))
                End If
            Else
                ZDate = "To_Date('" & Format(varValue, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
            End If
        Else
            If CStr(varValue) = "" Then
                'varValueΪ��
                If IsDate(varDefault) Then
                    If DateDiff("s", varDefault, "1899-12-30") = 0 Then
                        'varDefaultΪ��
                        ZDate = "NULL"
                    Else
                        ZDate = "To_Date('" & Format(varDefault, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                    End If
                Else
                    ZDate = IIf(CStr(varDefault) = "", "NULL", CStr(varDefault))
                End If
            Else
                ZDate = CStr(varValue)
            End If
        End If
    Else
        If IsDate(varValue) Then
            If DateDiff("s", varValue, "1899-12-30") = 0 Then
                'varValueΪ��
                If IsDate(varDefault) Then
                    If DateDiff("s", varDefault, "1899-12-30") = 0 Then
                        'varDefaultΪ��
                        ZDate = ""
                    Else
                        ZDate = Format(varDefault, "yyyy-mm-dd hh:mm:ss")
                    End If
                Else
                    ZDate = CStr(varDefault)
                End If
            Else
                ZDate = Format(varValue, "yyyy-mm-dd hh:mm:ss")
            End If
        Else
            If CStr(varValue) = "" Then
                'varValueΪ��
                If IsDate(varDefault) Then
                    If DateDiff("s", varDefault, "1899-12-30") = 0 Then
                        'varDefaultΪ��
                        ZDate = ""
                    Else
                        ZDate = Format(varDefault, "yyyy-mm-dd hh:mm:ss")
                    End If
                Else
                    ZDate = CStr(varDefault)
                End If
            Else
                ZDate = CStr(varValue)
            End If
        End If
    End If
End Function


Public Function GetWorkTrueDate(ByVal dtStart As Date, ByVal dtCur As Date, _
    Optional ByVal blnNextDate As Boolean = True, Optional ByVal blnEqualNextDay As Boolean = True) As String
    '���ݿ�ʼʱ�䣬ȷ��ָ��ʱ������ڣ�����ǰʱ���Ƿ��һ����һ�죨ת��Ϊͬһ��Ƚ�ʱ�֣�
    '��Σ�
    '   dtCur - ת������
    '   dtStart - ��ʼʱ��
    '   dtCur - ��ǰʱ��
    '   blnNextDate - True��ʾ��һ�죬False��ʾ��һ��
    '   blnEqualNextDay - ʱ�����ʱ�Ƿ�Ϊ��һ��
    If blnNextDate Then
        If DateDiff("n", Format(dtStart, "hh:mm:ss"), Format(dtCur, "hh:mm:ss")) < 0 Then
            dtStart = DateAdd("d", 1, dtStart)
        End If
        
        If blnEqualNextDay Then
            If DateDiff("n", Format(dtStart, "hh:mm:ss"), Format(dtCur, "hh:mm:ss")) = 0 Then
                dtStart = DateAdd("d", 1, dtStart)
            End If
        End If
    Else
        If DateDiff("n", Format(dtStart, "hh:mm:ss"), Format(dtCur, "hh:mm:ss")) > 0 Then
            dtStart = DateAdd("d", -1, dtStart)
        End If
    End If
    
    GetWorkTrueDate = Format(dtStart, "yyyy-mm-dd ") & Format(dtCur, "hh:mm:ss")
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlNumStrSort(ByVal strNums As String, Optional ByVal blnDelRepeat As Boolean, _
    Optional blnAsc As Boolean = True, Optional ByVal strSplit As String = ",") As String
    '�������ַ����������
    '��Σ�
    '   blnDelRepeat - ȥ���ظ���
    '   blnAsc - �Ƿ���������
    '   strSplit - �ָ���
    '˵�������뷨����,Ĭ������
    Dim varData As Variant, arrNum() As Long
    Dim varTemp As Variant, strReturn As String
    Dim i As Long, j As Long, k As Long
    Dim blnFind As Boolean
    
    If InStr(strNums, strSplit) = 0 Then ZlNumStrSort = Val(strNums): Exit Function
    varData = Split(strNums, strSplit)
    ReDim arrNum(0)
    For i = 0 To UBound(varData)
        blnFind = False
        If i = 0 Then
            blnFind = True
            arrNum(0) = Val(varData(i))
        Else
            For j = 0 To UBound(arrNum)
                If blnAsc Then '����
                    If arrNum(j) > Val(varData(i)) Then blnFind = True
                Else '����
                    If arrNum(j) < Val(varData(i)) Then blnFind = True
                End If
                If blnFind Then
                    ReDim Preserve arrNum(UBound(arrNum) + 1)
                    For k = UBound(arrNum) - 1 To j Step -1
                        arrNum(k + 1) = arrNum(k)
                    Next
                    arrNum(j) = Val(varData(i))
                    Exit For
                End If
            Next
        End If
        If blnFind = False Then
            ReDim Preserve arrNum(UBound(arrNum) + 1)
            arrNum(UBound(arrNum)) = Val(varData(i))
        End If
    Next
    
    For i = 0 To UBound(arrNum)
        If blnDelRepeat And i > 0 Then 'ȥ��
            If arrNum(i) <> arrNum(i - 1) Then
                strReturn = strReturn & strSplit & arrNum(i)
            End If
        Else
            strReturn = strReturn & strSplit & arrNum(i)
        End If
    Next
    If strReturn <> "" Then strReturn = Mid(strReturn, Len(strSplit) + 1)
    
    ZlNumStrSort = strReturn
End Function

Public Sub SetEnabled(ByVal objControls As Object, ByVal blnEnabled As Boolean)
    '���ÿؼ�����״̬
    Dim i As Integer
    
    On Error Resume Next
    For i = 0 To objControls.Count - 1
        If UCase(objControls(i).Name) <> UCase("cmdHelp") _
            And UCase(objControls(i).Name) <> UCase("cmdOk") _
            And UCase(objControls(i).Name) <> UCase("cmdCancel") _
            And UCase(TypeName(objControls(i))) <> UCase("Label") _
            And UCase(TypeName(objControls(i))) <> UCase("Frame") _
            And UCase(TypeName(objControls(i))) <> UCase("TabControl") _
            And UCase(TypeName(objControls(i))) <> UCase("PictureBox") _
            And UCase(TypeName(objControls(i))) <> UCase("VSFlexGrid") Then
            objControls(i).Enabled = blnEnabled
        End If
    Next
End Sub

Public Function GetClientPoint(ByVal objFrm As Form) As POINTAPI
'��ȡ��ǰָ���Ӧ�ڿؼ��е�λ��
    Dim pRet As POINTAPI
    Dim lngReturn As Long
    
    pRet = zlControl.GetCursorPosition()
    pRet.X = objFrm.ScaleX(pRet.X, vbPixels, vbTwips)
    pRet.Y = objFrm.ScaleY(pRet.Y, vbPixels, vbTwips)
    GetClientPoint = pRet
End Function

Public Sub SetEnabledBackColor(ByVal objControls As Object)
    '���ÿؼ�����״̬�벻����״̬�ı�����ɫ
    Dim i As Integer
    
    On Error Resume Next
    For i = 0 To objControls.Count - 1
        If UCase(TypeName(objControls(i))) = UCase("TextBox") _
            Or UCase(TypeName(objControls(i))) = UCase("ComboBox") Then
            objControls(i).BackColor = IIf(objControls(i).Enabled, vbWindowBackground, vbButtonFace)
        End If
    Next
End Sub

Public Sub SetBackColor(ByVal objControls As Object, ByVal lngBackColor As OLE_COLOR)
    '���ÿؼ�������ɫ
    Dim i As Integer
    On Error Resume Next
    
    For i = 0 To objControls.Count - 1
        If UCase(TypeName(objControls(i))) <> UCase("VSFlexGrid") _
            And UCase(TypeName(objControls(i))) <> UCase("TextBox") _
            And UCase(TypeName(objControls(i))) <> UCase("ComboBox") Then
            objControls(i).BackColor = lngBackColor
        End If
    Next
End Sub


Public Function GetSelectedIndex(ByVal OptionButtons As Object) As Integer
    '��ȡ��ѡ��ť���ѡ���������
    Dim i As Integer
    
    For i = OptionButtons.LBound To OptionButtons.UBound
        If OptionButtons(i).Value Then
            GetSelectedIndex = i: Exit For
        End If
    Next
End Function

Public Sub SetReportControlBackColorAlternate(rptData As ReportControl, Optional CustomColor As OLE_COLOR = -1)
    '����ReportControl���н���ɫ
    Dim i As Long, ObjItem As ReportRecordItem
    Dim lngRowCount As Long '�����к�
    
    On Error Resume Next
    For i = 0 To rptData.Rows.Count - 1
        If rptData.Rows(i).GroupRow Then
            lngRowCount = 0
        Else
            For Each ObjItem In rptData.Rows(i).Record
                If lngRowCount Mod 2 = 0 Then
                    ObjItem.BackColor = rptData.PaintManager.BackColor
                Else
                    ObjItem.BackColor = IIf(CustomColor = -1, G_AlternateColor, CustomColor)
                End If
            Next
            lngRowCount = lngRowCount + 1
        End If
    Next
End Sub

Public Sub SetVsGridRowChangeBackColor(ByVal vsGrid As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngNewRow As Long, _
    ByVal lngoldCol As Long, ByVal lngNewCol As Long, Optional CustomColor As OLE_COLOR = -1, _
    Optional ByVal lngStartCol As Long = -1, Optional ByVal lngEndCol As Long = -1)
    '���ܣ����иı�ʱ,������ص���ɫ
    '��Σ�
    '   CustomColor - �Զ�����ɫ��ѡ������ɫ
    '   lngStartCol - ��ʼ��
    '   lngEndCol - ��ֹ��
    '˵����RowData�п��ܴ�����ɫֵ
    Dim lngColStart As Long, lngColEnd As Long
    Err = 0: On Error Resume Next
    With vsGrid
        If lngOldRow > .FixedRows - 1 Then
            lngColStart = IIf(lngStartCol = -1 Or .IsSubtotal(lngOldRow), .FixedCols, lngStartCol)
            lngColEnd = IIf(lngEndCol = -1 Or .IsSubtotal(lngOldRow), .Cols - 1, lngEndCol)
            .Cell(flexcpBackColor, lngOldRow, lngColStart, lngOldRow, lngColEnd) = IIf(Val(.RowData(lngOldRow)) = 0, .BackColor, Val(.RowData(lngOldRow)))
        End If
        If lngNewRow > .FixedRows - 1 Then
            lngColStart = IIf(lngStartCol = -1 Or .IsSubtotal(lngNewRow), .FixedCols, lngStartCol)
            lngColEnd = IIf(lngEndCol = -1 Or .IsSubtotal(lngNewRow), .Cols - 1, lngEndCol)
            .Cell(flexcpBackColor, lngNewRow, lngColStart, lngNewRow, lngColEnd) = IIf(CustomColor <> -1, CustomColor, -2147483635) '16772055
        End If
    End With
End Sub

Public Function GetVsfGridData(rptData As ReportControl, _
    Optional ByVal strHiddenCols As String) As VSFlexGrid
    '����:��ReportControlת��ΪVSFlexGrid
    '���:
    '   strHiddenCols ����������(������0��ʼ)����ʽ����1,��2,��3,...
    Set GetVsfGridData = frmClinicPlanTemp.GetVsfGrid(rptData, strHiddenCols)
End Function

Public Function GetFirstCommandBar(ByRef objControls As CommandBarControls) As Long
'���ܣ���ȡ��������ӡԤ����ť��ĵ�һ����ť��index
    Dim objControl As CommandBarControl, idx As Long
    
    For Each objControl In objControls
        If objControl.ID = conMenu_File_Preview Then
            idx = objControl.index + 1
        End If
    Next
    GetFirstCommandBar = idx
End Function

Public Function FindNodeByKey(ByVal objNodes As Nodes, ByVal strKey As String) As Node
    '���ܣ�����Keyֵ�������ڵ�
    Dim objNode As Node
    
    For Each objNode In objNodes
        If objNode.Key = strKey Then
            Set FindNodeByKey = objNode
            Exit For
        End If
    Next
End Function

Public Function CollExitsValue(ByVal coll As Collection, ByVal strKey As String) As Boolean
'���ܣ����ݹؼ����ж�Ԫ���Ƿ�����ڼ�����
    Dim blnExits As Boolean
    
    If coll Is Nothing Then Exit Function
    CollExitsValue = True
    Err = 0: On Error Resume Next
    blnExits = IsObject(coll(strKey))
    If Err <> 0 Then Err = 0: CollExitsValue = False
End Function

Public Function AddRange(objOld As Collection, objAdd As Collection) As Collection
    'ƴ�Ӽ���
    Dim i As Long
    Dim objCol As New Collection
    
    Err = 0: On Error GoTo Errhand:
    If objOld Is Nothing And objAdd Is Nothing Then Exit Function
    If Not objOld Is Nothing Then
        For i = 1 To objOld.Count
            objCol.Add objOld(i)
        Next
    End If
    If Not objAdd Is Nothing Then
        For i = 1 To objAdd.Count
            objOld.Add objAdd(i)
        Next
    End If
    Set AddRange = objOld
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CalculatTimeInterval(ByVal BytMode As Byte, ByVal bln��ſ��� As Boolean, _
    ByRef intInterval As Integer, ByVal lngMaxStartSN As Long, _
    ByVal dtStartTime As Date, ByVal dtEndTime As Date, Optional ByVal str��Ϣʱ�� As String, _
    Optional ByVal lngStartSN As Long = 1, Optional ByVal lng��Լ�� As Long, _
    Optional ByVal strOriginalStartTime As String) As Collection
    '���ܣ����ݼ��ʱ�����������������ʱ��
    '��Σ�
    '   bytMode - ����ģʽ��0-���ݴ��������������㣬1-���������ż���
    '   intInterval - ���ʱ��(����)��bytMode=0ʱ����
    '   lngMaxStartSN - ������
    '   dtStartTime - ��ʼʱ��
    '   dtStartTime - ��ֹʱ��
    '   intStartSN - ��ʼ��ţ����ڶ��ʱ���������
    '   dtOriginalStartTime ԭʼ��ʱ��εĿ�ʼʱ�䣬��Ҫ��������Ϣʱ�ε�ʱ��Σ�
    '               �ڽ��мӺ�ʱ������Ŀ�ʼʱ��dtStartTime����ʱ��εĿ�ʼʱ�䣬�����޷��ж���Ϣʱ���ǵ��ջ��ǵڶ���
    '���Σ�
    '   intInterval - ���ʱ��(����)��bytMode=1ʱ����
    '���أ�
    '   Array(���,��ʼʱ��,��ֹʱ��,ԤԼ����)
    Dim objCol As Collection, lngCurSN As Long
    Dim dtCur As Date, dtTemp As Date
    Dim lngCount As Long, lngOverplus As Long
    Dim dtStart As Date, dtEnd As Date, varTimes As Variant
    Dim i As Long, j As Long
    Dim dtCurStart As Date, dtCurEnd As Date
    Dim objColAll As Collection
    Dim dtOriginalStartTime As Date
    
    Set objColAll = New Collection
    If strOriginalStartTime = "" Then
        dtOriginalStartTime = dtStartTime
    Else
        dtOriginalStartTime = CDate(strOriginalStartTime)
    End If
    If BytMode = 0 Then
        If str��Ϣʱ�� = "" Then
            dtCur = dtStartTime
            lngCurSN = lngStartSN
            Do While True
                If bln��ſ��� And lngCurSN > lngMaxStartSN Then Exit Do '����������ʱ����
                If DateDiff("n", dtEndTime, dtCur) >= 0 Then Exit Do '���ڵ�����ֹʱ��ʱ����
                dtTemp = DateAdd("n", intInterval, dtCur) '�Ӽ��ʱ��
                If DateDiff("n", dtTemp, dtEndTime) <= 0 Then
                    objColAll.Add Array(lngCurSN, dtCur, dtEndTime, IIf(lng��Լ�� > 0, 1, 0)), "K_" & lngCurSN
                    Exit Do '���ڵ�����ֹʱ��ʱ����
                Else
                    objColAll.Add Array(lngCurSN, dtCur, dtTemp, IIf(lng��Լ�� > 0, 1, 0)), "K_" & lngCurSN
                End If
                lngCurSN = lngCurSN + 1 '��ż�1
                lng��Լ�� = lng��Լ�� - 1
                dtCur = dtTemp
            Loop
        Else
            varTimes = Split(str��Ϣʱ��, ";")
            dtCurStart = dtStartTime
            For i = 0 To UBound(varTimes)
                '�����Ϣʱ�εĿ�ʼʱ��С�ڵ�ǰʱ�εĿ�ʼʱ�䣬���ʾ�ǵڶ��죬��Ϣʱ�εĿ�ʼʱ�����ֹʱ�䶼Ҫ��һ��
                dtStart = CDate(Format(dtOriginalStartTime, "yyyy-mm-dd ") & Split(varTimes(i), "-")(0))
                dtEnd = CDate(Format(dtOriginalStartTime, "yyyy-mm-dd ") & Split(varTimes(i), "-")(1))
                If DateDiff("n", dtStart, dtOriginalStartTime) > 0 Then dtStart = DateAdd("d", 1, dtStart): dtEnd = DateAdd("d", 1, dtEnd)
                '��Ϣʱ�ε���ֹʱ��С����Ϣʱ�εĿ�ʼʱ�䣬����Ϣʱ�ε���ֹʱ���һ��
                If DateDiff("n", dtEnd, dtStart) > 0 Then dtEnd = DateAdd("d", 1, dtEnd)
                
                If DateDiff("n", dtCurStart, dtStart) > 0 Then '��Ϣʱ�εĿ�ʼʱ����ڿ�ʼʱ�����Ч
                    dtCurEnd = dtStart '��ǰʱ�ε���ֹʱ�������Ϣʱ�εĿ�ʼʱ��
                    Set objCol = CalculatTimeInterval(0, bln��ſ���, intInterval, lngMaxStartSN, dtCurStart, dtCurEnd, "", _
                        lngStartSN + objColAll.Count, lng��Լ�� - objColAll.Count)
                    Set objColAll = AddRange(objColAll, objCol)
                    
                    '��һ��ʱ�εĿ�ʼʱ����ڵ�ǰ��Ϣʱ�ε���ֹʱ��
                    dtCurStart = dtEnd
                End If
            Next
            dtCurEnd = dtEndTime
            Set objCol = CalculatTimeInterval(0, bln��ſ���, intInterval, lngMaxStartSN, dtCurStart, dtCurEnd, "", _
                lngStartSN + objColAll.Count, lng��Լ�� - objColAll.Count)
            Set objColAll = AddRange(objColAll, objCol)
        End If
    ElseIf BytMode = 1 Then
        lngCount = GetMinuteCount(dtStartTime, dtEndTime, str��Ϣʱ��)
        intInterval = lngCount \ lngMaxStartSN 'ÿ��ʱ�ε�ƽ��������
        lngOverplus = lngCount - intInterval * lngMaxStartSN 'ʣ��δ������ķ������������䵽����������
        
        '���ʱ����������
        If intInterval = 0 Then lngMaxStartSN = lngOverplus
        
        If str��Ϣʱ�� = "" Then
            dtCur = dtStartTime
            For lngCurSN = 1 To lngMaxStartSN
                dtTemp = DateAdd("n", intInterval, dtCur) '�Ӽ��ʱ��
                If lngCurSN > lngMaxStartSN - lngOverplus Then
                    dtTemp = DateAdd("n", 1, dtTemp) '����δ�������ʱ�䣬�ڼ��ʱ��������ټ�1
                End If
                If DateDiff("n", dtTemp, dtEndTime) <= 0 Then
                    objColAll.Add Array(lngCurSN, dtCur, dtEndTime, IIf(lng��Լ�� > 0, 1, 0)), "K_" & lngCurSN
                    Exit For
                Else
                    objColAll.Add Array(lngCurSN, dtCur, dtTemp, IIf(lng��Լ�� > 0, 1, 0)), "K" & lngCurSN
                End If
                lng��Լ�� = lng��Լ�� - 1
                dtCur = dtTemp
            Next
        Else
            lngCurSN = 1
            varTimes = Split(str��Ϣʱ��, ";")
            dtCurStart = dtStartTime: lngCount = 0
            For j = 0 To UBound(varTimes)
                '�����Ϣʱ�εĿ�ʼʱ��С�ڵ�ǰʱ�εĿ�ʼʱ�䣬���ʾ�ǵڶ��죬��Ϣʱ�εĿ�ʼʱ�����ֹʱ�䶼Ҫ��һ��
                dtStart = CDate(Format(dtOriginalStartTime, "yyyy-mm-dd ") & Split(varTimes(j), "-")(0))
                dtEnd = CDate(Format(dtOriginalStartTime, "yyyy-mm-dd ") & Split(varTimes(j), "-")(1))
                If DateDiff("n", dtStart, dtOriginalStartTime) > 0 Then dtStart = DateAdd("d", 1, dtStart): dtEnd = DateAdd("d", 1, dtEnd)
                '��Ϣʱ�ε���ֹʱ��С����Ϣʱ�εĿ�ʼʱ�䣬����Ϣʱ�ε���ֹʱ���һ��
                If DateDiff("n", dtEnd, dtStart) > 0 Then dtEnd = DateAdd("d", 1, dtEnd)
                
                If DateDiff("n", dtCurStart, dtStart) > 0 Then '��Ϣʱ�εĿ�ʼʱ����ڿ�ʼʱ�����Ч
                    dtCurEnd = dtStart '��ǰʱ�ε���ֹʱ�������Ϣʱ�εĿ�ʼʱ��
                    For i = lngCurSN To lngMaxStartSN
                        dtCur = DateAdd("n", intInterval, dtCurStart) '�Ӽ��ʱ��
                        If i > lngMaxStartSN - lngOverplus Then
                            dtCur = DateAdd("n", 1, dtCur) '����δ�������ʱ�䣬�ڼ��ʱ��������ټ�1
                        End If
                        '��ǰʱ�ε���ֹʱ����ڵ�����Ϣʱ�εĿ�ʼʱ�䣬��ǰʱ�ε���ֹʱ�������Ϣʱ�εĿ�ʼʱ��
                        If DateDiff("n", dtCur, dtCurEnd) <= 0 Then
                            objColAll.Add Array(lngCurSN, dtCurStart, dtCurEnd, IIf(lng��Լ�� > 0, 1, 0)), "K" & i
                            lngCurSN = lngCurSN + 1
                            lng��Լ�� = lng��Լ�� - 1
                            dtCurStart = dtEnd
                            Exit For
                        Else
                            objColAll.Add Array(lngCurSN, dtCurStart, dtCur, IIf(lng��Լ�� > 0, 1, 0)), "K" & i
                            lngCurSN = lngCurSN + 1
                            lng��Լ�� = lng��Լ�� - 1
                            dtCurStart = dtCur
                        End If
                    Next
                End If
            Next
            dtCurEnd = dtEndTime
            For i = lngCurSN To lngMaxStartSN
                dtCur = DateAdd("n", intInterval, dtCurStart) '�Ӽ��ʱ��
                If i > lngMaxStartSN - lngOverplus Then
                    dtCur = DateAdd("n", 1, dtCur) '����δ�������ʱ�䣬�ڼ��ʱ��������ټ�1
                End If
                '��ǰʱ�ε���ֹʱ����ڵ�����Ϣʱ�εĿ�ʼʱ�䣬��ǰʱ�ε���ֹʱ�������Ϣʱ�εĿ�ʼʱ��
                If DateDiff("n", dtCur, dtCurEnd) <= 0 Then
                    objColAll.Add Array(lngCurSN, dtCurStart, dtCurEnd, IIf(lng��Լ�� > 0, 1, 0)), "K" & i
                    lng��Լ�� = lng��Լ�� - 1
                    Exit For
                Else
                    objColAll.Add Array(lngCurSN, dtCurStart, dtCur, IIf(lng��Լ�� > 0, 1, 0)), "K" & i
                    lngCurSN = lngCurSN + 1
                    lng��Լ�� = lng��Լ�� - 1
                    dtCurStart = dtCur
                End If
            Next
        End If
    End If
    Set CalculatTimeInterval = objColAll
End Function


Public Function GetMinuteCount(ByVal dtStartTime As Date, ByVal dtEndTime As Date, ByVal str��Ϣʱ�� As String) As Long
    '��ȡʱ�ε��ܵķ�����
    Dim dtCurStart As Date, dtCurEnd As Date
    Dim lngCount As Long, varTimes As Variant
    Dim i As Long, dtStart As Date, dtEnd As Date
    
    Err = 0: On Error GoTo Errhand:
    dtStartTime = Format(dtStartTime, "yyyy-mm-dd") & " " & Format(dtStartTime, "HH:MM")
    dtEndTime = GetWorkTrueDate(dtStartTime, dtEndTime)
    
    If str��Ϣʱ�� = "" Then
        lngCount = DateDiff("n", dtStartTime, dtEndTime) '�ܵķ�����
    Else
        varTimes = Split(str��Ϣʱ��, ";")
        dtCurStart = dtStartTime: lngCount = 0
        For i = 0 To UBound(varTimes)
            '�����Ϣʱ�εĿ�ʼʱ��С�ڵ�ǰʱ�εĿ�ʼʱ�䣬���ʾ�ǵڶ��죬��Ϣʱ�εĿ�ʼʱ�����ֹʱ�䶼Ҫ��һ��
            dtStart = CDate(Format(dtCurStart, "yyyy-mm-dd ") & Split(varTimes(i), "-")(0))
            dtEnd = CDate(Format(dtCurStart, "yyyy-mm-dd ") & Split(varTimes(i), "-")(1))
            If DateDiff("n", dtStart, dtCurStart) > 0 Then dtStart = DateAdd("d", 1, dtStart): dtEnd = DateAdd("d", 1, dtEnd)
            '��Ϣʱ�ε���ֹʱ��С����Ϣʱ�εĿ�ʼʱ�䣬����Ϣʱ�ε���ֹʱ���һ��
            If DateDiff("n", dtEnd, dtStart) > 0 Then dtEnd = DateAdd("d", 1, dtEnd)
            
            dtCurEnd = dtStart '��ǰʱ�ε���ֹʱ�������Ϣʱ�εĿ�ʼʱ��
            lngCount = lngCount + DateDiff("n", dtCurStart, dtCurEnd) '�ܵķ�����
            
            '��һ��ʱ�εĿ�ʼʱ����ڵ�ǰ��Ϣʱ�ε���ֹʱ��
            dtCurStart = dtEnd
        Next
        dtCurEnd = dtEndTime
        lngCount = lngCount + DateDiff("n", dtCurStart, dtCurEnd) '�ܵķ�����
    End If
    GetMinuteCount = lngCount
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetMaxLocalCode(ByVal strTableName As String, _
    Optional ByVal strColumn As String = "����", Optional ByVal strWhere As String) As String
    '���ܣ���ȡָ�����������(���漰�ϼ�����)
    '��Σ�����
    '���Σ��ɹ����� ������; ���߷��� ��
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim strNextCode As String
    
    Err = 0: On Error GoTo ErrHandler
    '�����볤��
    strSQL = "Select Nvl(Max(Length(" & strColumn & ")), 1) CodeLen" & vbNewLine & _
            " From " & strTableName & " Where 1=1 " & strWhere
    '������
    strSQL = "Select Max(LPad(" & strColumn & ", CodeLen, '0')) As MaxCode" & vbNewLine & _
            " From " & strTableName & " A, (" & strSQL & ") B" & vbNewLine & _
            " Where 1=1 " & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetMaxLocalCode")
    If rsTemp.EOF Then GetMaxLocalCode = "1": Exit Function
    
    strNextCode = zlStr.Increase(Nvl(rsTemp!MaxCode, "0")) '��1
    If strNextCode = "" Then strNextCode = "1"
    
    GetMaxLocalCode = strNextCode
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    GetMaxLocalCode = ""
End Function

Public Function GetWeekCount(ByVal intYear As Integer, ByVal intMonth As Integer) As Integer
    '���ܣ�����ĳ��ĳ��һ���ж�����
    Dim dtCur As Date, intCount As Integer
    
    dtCur = CDate(intYear & "-" & intMonth & "-01")
    Do While Month(dtCur) = intMonth
        intCount = intCount + 1
        dtCur = DateAdd("ww", 1, dtCur)
    Loop
    '������С�����ڣ����ټ�һ��
    'Weekday(dtStartDate, vbMonday):1-����һ,2-���ڶ�,3-������,4-������,5-������,6-������,7-������
    If Day(dtCur) < Weekday(dtCur, vbMonday) Then intCount = intCount + 1
    GetWeekCount = intCount
End Function

Public Function GetDateWeek(ByVal dtStart As Date) As Integer
    '���ܣ����ݵ�ǰ���ڼ����ǵڼ���
    Dim dtCur As Date, intWeek As Integer
    Dim dtFirstDayWeek As Integer
    
    dtCur = Format(dtStart, "yyyy-mm-01")
    dtFirstDayWeek = Weekday(dtCur, vbMonday) '��¼��һ�����ڼ�
    Do While Month(dtCur) = Month(dtStart) And DateDiff("d", dtCur, dtStart) > 0
        If Weekday(dtCur, vbMonday) = dtFirstDayWeek Then '���ڵ�һ����������һ��
            intWeek = intWeek + 1
        End If
        dtCur = DateAdd("d", 1, dtCur)
    Loop
    '��ǰ���ڵ�����С�ڵ��ڱ��µ�һ������ڣ����ټ�һ��
    If Weekday(dtCur, vbMonday) <= dtFirstDayWeek Then intWeek = intWeek + 1
    GetDateWeek = intWeek
End Function

Public Function GetDateRange(ByVal intYear As Integer, ByVal intMonth As Integer, Optional ByVal intWeek As Integer) As Variant
    '���ܣ�������������Ϣ��ȡ���ڷ�Χ
    '���أ����飬0-��ʼʱ�䣬1-��ֹʱ��
    Dim datStart As Date, datEnd As Date, datFirstDay As Date
    
    datFirstDay = intYear & "-" & intMonth & "-01"
    If intWeek = 0 Then
        '���µ�һ���һ���ټ�һ����Ǳ��µĿ�ʼ���ںͽ�������
        datStart = datFirstDay
        datEnd = DateAdd("d", -1, DateAdd("m", 1, datStart))
    Else
        datStart = datFirstDay
        datEnd = datFirstDay
        '�����һ�������һ���ͼ�һ��
        If Weekday(datEnd) = vbMonday Then datEnd = DateAdd("d", 1, datEnd)
        Do While Weekday(datEnd, vbMonday) < 7
            datEnd = DateAdd("d", 1, datEnd)
        Loop
        If intWeek > 1 Then
            datStart = DateAdd("ww", intWeek - 2, DateAdd("d", 1, datEnd))
            datEnd = DateAdd("ww", intWeek - 1, datEnd)
            If DateDiff("m", datFirstDay, datEnd) > 0 Then '������,ֱ��ȡ���һ��
                datEnd = DateAdd("d", -1, DateAdd("m", 1, datFirstDay))
            End If
        End If
    End If
    '��ʼʱ��С�ڵ�ǰ���ڣ�������쿪ʼ
    'If DateDiff("d", datStart, CDate(Format(Now, "yyyy-mm-dd"))) >= 0 Then datStart = CDate(Format(DateAdd("d", Now, 1), "yyyy-mm-dd"))
    GetDateRange = Array(datStart, datEnd)
End Function

Public Function GetWeekIndex(ByVal strWeek As String) As Integer
    '�����������ƻ�ȡ����
    Select Case strWeek
    Case "��һ", "����һ"
        GetWeekIndex = 0
    Case "�ܶ�", "���ڶ�"
        GetWeekIndex = 1
    Case "����", "������"
        GetWeekIndex = 2
    Case "����", "������"
        GetWeekIndex = 3
    Case "����", "������"
        GetWeekIndex = 4
    Case "����", "������"
        GetWeekIndex = 5
    Case "����", "������"
        GetWeekIndex = 6
    Case Else
        GetWeekIndex = -1
    End Select
End Function

Public Function GetWeekName(ByVal intIndex As Integer, Optional ByVal blnShort As Boolean = True) As String
    '����������ȡ��������
    '   blnShort True-��"��һ"��False-��"����һ"
    Select Case intIndex
    Case 0
        GetWeekName = IIf(blnShort, "��һ", "����һ")
    Case 1
        GetWeekName = IIf(blnShort, "�ܶ�", "���ڶ�")
    Case 2
        GetWeekName = IIf(blnShort, "����", "������")
    Case 3
        GetWeekName = IIf(blnShort, "����", "������")
    Case 4
        GetWeekName = IIf(blnShort, "����", "������")
    Case 5
        GetWeekName = IIf(blnShort, "����", "������")
    Case 6
        GetWeekName = IIf(blnShort, "����", "������")
    Case Else
        GetWeekName = ""
    End Select
End Function

Public Function GetTempImage(ByVal strTxt As String, _
    Optional ByVal dblWidth As Double, Optional ByVal dblHight As Double, _
    Optional ByVal lngBackColor As OLE_COLOR = vbButtonFace, _
    Optional ByVal lngForeColor As OLE_COLOR = vbBlue, _
    Optional ByVal objFont As StdFont, _
    Optional ByVal intAlignment As PictureTextAlignmentSettings = pictxtAlignCenterCenter, _
    Optional ByVal strSubTxt As String, _
    Optional ByVal lngSubForeColor As OLE_COLOR = vbBlack, _
    Optional ByVal objSubFont As StdFont, _
    Optional ByVal intSubAlignment As PictureTextAlignmentSettings = pictxtAlignCenterCenter) As IPictureDisp
    '���ܣ����ݲ�������ͼƬ
    '��Σ�
    '   strTxt - ��Ҫ��ʾ�ı�
    '   dblWidth,dblHight - ͼƬ��С��ȱʡΪ�ı���ӡ������Ŀ�Ⱥ͸߶�
    '   lngBackColor - ͼƬ����ɫ��ȱʡΪ��ť������ɫ
    '   lngForeColor - ��Ҫ�ı���ǰ��ɫ
    '   objFont - ��Ҫ�ı�������
    '   intAlignment - ��Ҫ�ı������λ��
    '   strSubTxt - �����ı�
    '   lngSubForeColor - �����ı���ǰ��ɫ
    '   objSubFont - �����ı�������
    '   intSubAlignment - �����ı������λ��
    '���أ�ͼƬ����
    Set GetTempImage = frmClinicPlanTemp.GetTempPicture(strTxt, dblWidth, dblHight, lngBackColor, _
        lngForeColor, objFont, intAlignment, strSubTxt, lngSubForeColor, objSubFont, intSubAlignment)
End Function

Public Function FormatApplyToStr(ByVal strApply As String) As String
    '��ʽ����Ӧ�����ַ���
    '��Σ�
    '   strApply yyyy-mm-dd|yyyy-mm-dd|yyyy-mm-dd|...
    '˵�������ڰ��찲�ų������ģ��
    Dim varStr As Variant, i As Integer
    Dim strNewStr As String
    
    If InStr(strApply, "|") = 0 And IsDate(strApply) = False Then
        FormatApplyToStr = strApply
        Exit Function
    End If
    
    varStr = Split(strApply, "|")
    For i = 0 To UBound(varStr)
        strNewStr = strNewStr & "|" & Day(varStr(i)) & "��"
    Next
    If strNewStr <> "" Then strNewStr = Mid(strNewStr, 2)
    FormatApplyToStr = strNewStr
End Function

Public Function CheckTimeBucketIsCross(ByVal dtStartA As Date, ByVal dtEndA As Date, _
    ByVal dtStartB As Date, ByVal dtEndB) As Boolean
    '���ʱ����Ƿ��н���
    '˵����
    '   1.���ʱ��ο����죬����Ҫ�ٷֶμ��
    '   2.����ʱ��������Ѵ���Ϊ�ο�ͬһ���
    Dim blnTwoDayA As Boolean, blnTwoDayB As Boolean
    Dim blnHaveCross As Boolean
    Dim dtStartTemp As Date, dtEndTemp As Date
    
    blnTwoDayA = DateDiff("d", dtStartA, dtEndA) > 0
    blnTwoDayB = DateDiff("d", dtStartB, dtEndB) > 0
    If blnTwoDayA And blnTwoDayB Then '����ʱ��ξ�����,�϶��н���
        CheckTimeBucketIsCross = True
        Exit Function
    End If
    
    blnHaveCross = Not (DateDiff("n", dtStartA, dtEndB) <= 0 Or DateDiff("n", dtEndA, dtStartB) >= 0)
    If blnHaveCross Then CheckTimeBucketIsCross = True: Exit Function '��֪�н��棬ֱ���˳�
    
    If blnTwoDayA And blnTwoDayB = False Then 'Aʱ��ο��죬Bʱ��β�����
        '��Aʱ��εڶ���Ĳ�����B�Ƚ�
        dtStartTemp = CDate(Format(dtStartA, "yyyy-mm-dd 00:00"))
        dtEndTemp = CDate(Format(dtStartA, "yyyy-mm-dd ") & Format(dtEndA, "HH:mm"))
        blnHaveCross = Not (DateDiff("n", dtStartTemp, dtEndB) <= 0 Or DateDiff("n", dtEndTemp, dtStartB) >= 0)
    ElseIf blnTwoDayA = False And blnTwoDayB Then 'Aʱ��β����죬Bʱ��ο���
        '��Bʱ��εڶ���Ĳ�����A�Ƚ�
        dtStartTemp = CDate(Format(dtStartB, "yyyy-mm-dd 00:00"))
        dtEndTemp = CDate(Format(dtStartB, "yyyy-mm-dd ") & Format(dtEndB, "HH:mm"))
        blnHaveCross = Not (DateDiff("n", dtStartA, dtEndTemp) <= 0 Or DateDiff("n", dtEndA, dtStartTemp) >= 0)
    Else '����ʱ��ξ������죬�ѱȽϳ����
        'blnHaveCross
    End If
    CheckTimeBucketIsCross = blnHaveCross
End Function

Public Function IsDoubleMonthWeekPlan(ByRef intYear As Integer, ByRef intMonth As Integer, _
    ByRef intWeek As Integer, ByRef dtStartDate As Date, ByRef dtEndDate As Date) As Boolean
    '�жϲ���ȡ���µ��ܰ��ŵ���һ��������������
    '��Σ�
    '   dtStartDate��dtEndDate Ҫ�жϵĳ�����ʱ�䷶Χ
    '���أ�������ڿ����򷵻�True�����򷵻�False
    '˵�����������True����
    '       intWeek������һ��������������
    '       dtStartDate��dtEndDate�ֱ𷵻�������(����)�Ŀ�ʼʱ��ͽ���ʱ��
    
    If DateDiff("d", dtStartDate, dtEndDate) >= 6 Then Exit Function
    
    '���ڿ��µģ�������һ��������������
    If Month(DateAdd("d", -1, dtStartDate)) <> Month(dtStartDate) Then
        '��ǰ�ǵ�һ��
        dtStartDate = DateAdd("d", DateDiff("d", dtStartDate, dtEndDate) - 6, dtStartDate)
        intYear = Year(dtStartDate): intMonth = Month(dtStartDate)
        intWeek = GetWeekCount(intYear, intMonth)
    ElseIf Month(DateAdd("d", 1, dtEndDate)) <> Month(dtEndDate) Then
        '��ǰ�����һ��
        dtEndDate = DateAdd("d", 6 - DateDiff("d", dtStartDate, dtEndDate), dtEndDate)
        intYear = Year(dtEndDate): intMonth = Month(dtEndDate)
        intWeek = 1
    End If
    IsDoubleMonthWeekPlan = True
End Function

Public Function GetPopupCommandBar(frmMain As Form, cbsMain As CommandBars, _
    Optional ByVal lngControlPopupID As Long = conMenu_EditPopup) As CommandBar
    '���������˵�
    Dim objPopup As CommandBarPopup, cbCommandBar As CommandBar
    Dim cbrControl As CommandBarControl, cbrControlNew As CommandBarControl
    Dim i As Integer
    
    Set objPopup = cbsMain.FindControl(xtpControlPopup, lngControlPopupID, , True)
    If objPopup Is Nothing Then Exit Function
    Set cbCommandBar = cbsMain.Add("Popup", xtpBarPopup) '�����˵�
    If cbCommandBar Is Nothing Then Exit Function
    
    For i = 1 To objPopup.CommandBar.Controls.Count
        Set cbrControl = objPopup.CommandBar.Controls(i)
        Call frmMain.zlUpdateCommandBars(cbrControl) '�ж��Ƿ�ɼ�����Ϊ��һ��ʱ�˵���û��ִ��Update
        If cbrControl.Visible Then
            Set cbrControlNew = cbCommandBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
            cbrControlNew.BeginGroup = cbrControl.BeginGroup
            cbrControlNew.Enabled = cbrControl.Enabled
        End If
    Next
    Set GetPopupCommandBar = cbCommandBar
End Function

