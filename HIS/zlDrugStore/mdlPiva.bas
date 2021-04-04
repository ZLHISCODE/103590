Attribute VB_Name = "mdlPiva"
Option Explicit

Private mrsTrans As ADODB.Recordset             '��Һ����¼��������Һ�����ݣ�ҩƷ��
Private mrsPRI As Recordset                     '��ҺҩƷ���ȼ�
Private mrsVol As Recordset                     '������������
Private mrstemp As Recordset                    '��ʱ��¼
Private mblnLastBatch As Boolean                '�Ƿ񱣳��ϴ�����
Private Sub Piva_GetPara()
    'ȡ��������һЩ����
    If mrsPRI Is Nothing Then
        gstrSQL = "select ����id,��������,��ҩ����,Ƶ��,��Ч,���ȼ� from ��ҺҩƷ���ȼ� order by ���ȼ�"
        Set mrsPRI = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ȼ�����")
    End If
    
    If mrsVol Is Nothing Then
        gstrSQL = "select ����id,��������,����,��ҩ���� from ������������"
        Set mrsVol = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������������")
    End If
    
    mblnLastBatch = (Val(zlDatabase.GetPara("�����ϴ�����", glngSys, 1345, 0)) = 1)
End Sub


Public Function PIVA_AutoSetBatch(ByVal lng�ⷿid As Long, ByVal strSendNO As String) As Boolean
    '�Զ���������
    'lng�ⷿid���������Ĳ���id
    'strdSendNO��ҽ�����ͺ�
    Dim rsTrans As ADODB.Recordset
    Dim lng��ҩid As Long
    Dim int��ҩ���� As Integer
    Dim rstemp As Recordset
    Dim strOldִ��ʱ�� As String
    Dim strOld���� As String
    Dim lngOld����id As Long
    Dim lngOld��ҩid As Long
    Dim lng���� As Long
    Dim lngCount As Long
    Dim lng���ȼ� As Long
    Dim int��� As Integer
    Dim i As Integer
    Dim strInput As String
    Dim arrExecute As Variant
    
    On Error GoTo errHandle
    
    Call Piva_GetPara
    Call Piva_IniTransRec
    Call Piva_IniPriRec
    
    Set rsTrans = Piva_GetTrans(lng�ⷿid, strSendNO)
    
    With rsTrans
        Set rstemp = rsTrans
        If .RecordCount > 0 Then
            rsTrans.Sort = "����,����id,��ҩid,ִ��ʱ��,����"
        End If
        Do While Not .EOF
            lngCount = lngCount + 1
            
            '�������ȼ������������������κ����ȼ�
            If mrsPRI.RecordCount > 0 Or mrsVol.RecordCount > 0 Then
                
                If strOldִ��ʱ�� <> IIf(IsNull(!ִ��ʱ��), "", Format(!ִ��ʱ��, "YYYY-MM-DD")) Then
                    '������һ�����˵�����
                    If lngCount > 1 Then
                        If mrsPRI.RecordCount > 0 Then
                            Call Piva_Set���ȼ�(mrstemp, mrsTrans, lngOld��ҩid)
                        End If
                        
                        If mrsVol.RecordCount > 0 Then
                            Call Piva_Set����(mrsTrans, lngOld����id, strOld����, strOldִ��ʱ��)
                        End If
                    End If
                    
                    '��ǰ���˵�����
                    Call Piva_IniPriRec
                    strOldִ��ʱ�� = IIf(IsNull(!ִ��ʱ��), "", Format(!ִ��ʱ��, "YYYY-MM-DD"))
                    strOld���� = IIf(IsNull(!��ҩ����), "", !��ҩ���� & "#")
                    lngOld����id = IIf(IsNull(!����ID), 0, !����ID)
                    lngOld��ҩid = !��ҩid
                Else
                    If strOld���� <> IIf(IsNull(!��ҩ����), "", !��ҩ���� & "#") Then
                        '������һ�����˵�����
                        If mrsPRI.RecordCount > 0 Then
                            Call Piva_Set���ȼ�(mrstemp, mrsTrans, lngOld��ҩid)
                        End If
                        
                        If mrsVol.RecordCount > 0 Then
                            Call Piva_Set����(mrsTrans, lngOld����id, strOld����, strOldִ��ʱ��)
                        End If
                        '��ǰ���˵�����
                        Call Piva_IniPriRec
                        strOld���� = IIf(IsNull(!��ҩ����), "", !��ҩ���� & "#")
                        lngOld����id = IIf(IsNull(!����ID), 0, !����ID)
                        lngOld��ҩid = !��ҩid
                        
                    Else
                        If lngOld����id <> IIf(IsNull(!����ID), 0, !����ID) Then
                            '������һ�����˵�����
                            If mrsPRI.RecordCount > 0 Then
                                Call Piva_Set���ȼ�(mrstemp, mrsTrans, lngOld��ҩid)
                            End If
                            
                            If mrsVol.RecordCount > 0 Then
                                Call Piva_Set����(mrsTrans, lngOld����id, strOld����, strOldִ��ʱ��)
                            End If
                            '��ǰ���˵�����
                            Call Piva_IniPriRec
                            lngOld����id = IIf(IsNull(!����ID), 0, !����ID)
                            lngOld��ҩid = !��ҩid
                        Else
                            If lngOld��ҩid <> !��ҩid Then
                                If Not mrsPRI.RecordCount Then
                                    Call Piva_Set���ȼ�(mrstemp, mrsTrans, lngOld��ҩid)
                                End If
                                lngOld��ҩid = !��ҩid
                            End If
                        End If
                    End If
                End If
                
                '�������ݼ�
                mrstemp.AddNew
                mrstemp!����ID = !���˿���id
                mrstemp!��ҩid = !��ҩid
                mrstemp!��ҩ���� = IIf(IsNull(!��ҩ����), "", !��ҩ����)
                mrstemp!Ƶ�� = IIf(IsNull(!ִ��Ƶ��), "", !ִ��Ƶ��)
                mrstemp.Update
            End If
             
            mrsTrans.AddNew
            mrsTrans!��ҩid = !��ҩid
            mrsTrans!����ID = !����ID
            mrsTrans!��� = !���
            mrsTrans!���� = IIf(IsNull(!����), "", !����)
            mrsTrans!�Ա� = IIf(IsNull(!�Ա�), "", !�Ա�)
            mrsTrans!���� = IIf(IsNull(!����), "", !����)
            mrsTrans!סԺ�� = IIf(IsNull(!סԺ��), "", !סԺ��)
            mrsTrans!���� = IIf(IsNull(!����), "", !����)
            mrsTrans!���˲��� = !���˲���
            mrsTrans!���˿��� = !���˿���
            mrsTrans!ִ��ʱ�� = IIf(IsNull(!ִ��ʱ��), "", Format(!ִ��ʱ��, "YYYY-MM-DD HH:MM:SS"))
            mrsTrans!����ID = IIf(IsNull(!����ID), 0, !����ID)
            mrsTrans!��ҳid = IIf(IsNull(!��ҳid), 0, !��ҳid)
            mrsTrans!���˿���id = IIf(IsNull(!���˿���id), 0, !���˿���id)
            mrsTrans!���ʱ�� = IIf(IsNull(!���ʱ��), "", !���ʱ��)
            
            mrsTrans!��ҩ���� = IIf(IsNull(!��ҩ����), "", !��ҩ���� & "#")
            mrsTrans!����ҩ���� = IIf(IsNull(!��ҩ����), "", !��ҩ���� & "#")
            mrsTrans!ƿǩ�� = IIf(IsNull(!ƿǩ��), "", !ƿǩ��)
            mrsTrans!��ӡ��־ = IIf(IIf(IsNull(!��ӡ��־), 0, !��ӡ��־) = 0, 0, 1)
            mrsTrans!�Ƿ��� = IIf(IsNull(!�Ƿ���), 0, !�Ƿ���)
            mrsTrans!�˲��� = IIf(IsNull(!������Ա), "", !������Ա)
            mrsTrans!�˲�ʱ�� = IIf(IsNull(!����ʱ��), "", Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS"))
            mrsTrans!��ҩ�� = IIf(IsNull(!������Ա), "", !������Ա)
            mrsTrans!��ҩʱ�� = IIf(IsNull(!����ʱ��), "", Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS"))
            mrsTrans!��ҩ���� = IIf(IsNull(!��ҩ����), "", !��ҩ����)
            mrsTrans!��ҩ�� = IIf(IsNull(!������Ա), "", !������Ա)
            mrsTrans!��ҩʱ�� = IIf(IsNull(!����ʱ��), "", Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS"))
            mrsTrans!������ = IIf(IsNull(!������Ա), "", !������Ա)
            mrsTrans!����ʱ�� = IIf(IsNull(!����ʱ��), "", Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS"))
            mrsTrans!���������� = IIf(IsNull(!������Ա), "", !������Ա)
            mrsTrans!��������ʱ�� = IIf(IsNull(!����ʱ��), "", Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS"))
            mrsTrans!��������� = IIf(IsNull(!������Ա), "", !������Ա)
            mrsTrans!�������ʱ�� = IIf(IsNull(!����ʱ��), "", Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS"))
            mrsTrans!����ҩ�� = 1
            mrsTrans!ҩʦ���ʱ�� = IIf(IsNull(!ҩʦ���ʱ��), 0, !ҩʦ���ʱ��)
            mrsTrans!�Ƿ�������� = IIf(IsNull(!�Ƿ��������), 0, !�Ƿ��������)
            mrsTrans!�Ƿ����� = IIf(IsNull(!�Ƿ�����), 0, !�Ƿ�����)
            mrsTrans!�ֹ��������� = IIf(IsNull(!�ֹ���������), 0, !�ֹ���������)
            mrsTrans!����ԭ�� = IIf(IsNull(!����ԭ��), "", !����ԭ��)
            
            mrsTrans!�շ�Id = !�շ�Id
            mrsTrans!���� = !����
            mrsTrans!NO = !NO
            mrsTrans!ҩƷ���� = "[" & !ҩƷ���� & "]" & !ͨ����
            mrsTrans!ͨ���� = !ͨ����
            mrsTrans!��Ʒ�� = IIf(IsNull(!��Ʒ��), "", !��Ʒ��)
            mrsTrans!Ӣ���� = IIf(IsNull(!Ӣ����), "", !Ӣ����)
            mrsTrans!��� = IIf(IsNull(!���), "", !���)
            mrsTrans!���� = IIf(IsNull(!����), "", !����)
            mrsTrans!���� = IIf(IsNull(!����), "", !����)
            mrsTrans!���� = IIf(IsNull(!����), 0, !����)
            mrsTrans!������λ = !������λ
            mrsTrans!Ƶ�� = IIf(IsNull(!Ƶ��), "", !Ƶ��)
            mrsTrans!���� = IIf(IsNull(!����), 0, !����)
            mrsTrans!��λ = !��λ
            mrsTrans!���� = !����
            mrsTrans!�÷� = IIf(IsNull(!�÷�), "", !�÷�)
            mrsTrans!ҩƷID = IIf(IsNull(!ҩƷID), 0, !ҩƷID)
            mrsTrans!ҩ��ID = !ҩ��ID
            mrsTrans!������� = !�������
            mrsTrans!����ID = !����ID
            mrsTrans!��ҩ���� = !��ҩ����
            
            mrsTrans!��ҩ���� = IIf(IsNull(!��ҩ����), 0, !��ҩ����)
            mrsTrans!������� = IIf(IsNull(!�������), 0, !�������)
            mrsTrans!ʵ������ = IIf(IsNull(!ʵ������), 0, !�������)
            
            mrsTrans!ҽ��id = !ҽ��id
            mrsTrans!���ͺ� = !���ͺ�
            mrsTrans!ҽ������ʱ�� = IIf(IsNull(!ҽ������ʱ��), "", Format(!ҽ������ʱ��, "YYYY-MM-DD HH:MM:SS"))
            mrsTrans!����� = IIf(IsNull(!�����), 0, !�����)
            mrsTrans!�������� = IIf(IsNull(!��������), "", !��������)
            mrsTrans!���� = !����
            mrsTrans!��ɫ = !��ɫ
            mrsTrans!ִ�б�־ = 0
            mrsTrans!��ý = IIf(IsNull(!��ý), 0, !��ý)
            
            If !��ҩid <> lng��ҩid Then
                int��� = int��� + 1
            End If
            mrsTrans!��� = int���
            mrsTrans.Update
            

            If lngCount = .RecordCount Then
                If mrsPRI.RecordCount > 0 Then
                    Call Piva_Set���ȼ�(mrstemp, mrsTrans, lngOld��ҩid)
                End If
                
                If mrsVol.RecordCount > 0 Then
                    Call Piva_Set����(mrsTrans, lngOld����id, strOld����, strOldִ��ʱ��)
                End If
            End If

            If !��ҩid = lng��ҩid Then
                If Val(!��ҩ����) > 0 Then
                    int��ҩ���� = 1
                ElseIf int��ҩ���� = 0 And Val(!��ҩ����) = 0 Then
                    int��ҩ���� = 0
                End If
            Else
                int��ҩ���� = Val(!��ҩ����)
            End If
            
            mrsTrans.Filter = "��ҩid=" & !��ҩid
            Do While Not mrsTrans.EOF
                mrsTrans.Update "����ҩ��", int��ҩ����
                mrsTrans.MoveNext
            Loop
            
            mrsTrans.Filter = ""
            lng��ҩid = !��ҩid
            
            .MoveNext
        Loop
    End With
    
    lng��ҩid = 0
    
    '�����ϴ�����
    If mblnLastBatch = True Then
        Call Piva_SetLastBatch(mrsTrans)
    End If
    
    With mrsTrans
        .Filter = ""
        .Sort = "��ҩID"
        Do While Not .EOF
            If !����ҩ���� <> !��ҩ���� And lng��ҩid <> Val(!��ҩid) Then
                lng��ҩid = Val(!��ҩid)
                
                If IIf(IsNull(!����ҩ����), "", !����ҩ����) = "" Then
                    strInput = IIf(strInput = "", "", strInput & "|") & !��ҩid & ",:" & IIf(IsNull(!���ȼ�), 0, !���ȼ�)
                Else
                    strInput = IIf(strInput = "", "", strInput & "|") & !��ҩid & "," & Mid(!����ҩ����, 1, IIf(Len(!����ҩ����) = 0, 0, Len(!����ҩ����) - 1)) & ":" & IIf(IsNull(!���ȼ�), 0, !���ȼ�)
                End If
            End If
            .MoveNext
        Loop
    End With
    
    If strInput <> "" Then
        arrExecute = Piva_GetArrayByStr(strInput, 3900, "|")
        For i = 0 To UBound(arrExecute)
            gstrSQL = "Zl_��Һ��ҩ��¼_����("
            '��ҩID,����
            gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
            gstrSQL = gstrSQL & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "Piva_AutoSetBatch")
        Next
    End If
    
    PIVA_AutoSetBatch = True
    
    Exit Function
errHandle:
    PIVA_AutoSetBatch = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Piva_GetArrayByStr(ByVal strInput As String, ByVal lngLength As Long, ByVal strSplitChar As String) As Variant
    '���ݴ�����ַ������зֽ⣬����ָ���ַ����Ⱦ���Ҫ���зֽ⣬������浽������
    '��Σ�strInput-������ַ�����strSplitChar-�ַ��������ݵķָ���
    '���أ����飬���������Ա���ַ����Ȳ�����ָ������
    Dim strArray As Variant
    Dim ArrTmp As Variant
    Dim strTmp As String
    Dim lngCount As Long
    Dim i As Long
    
    strArray = Array()
   
    '����ָ���ַ�ʱ����Ҫ�ֽ�
    If Len(strInput) > lngLength Then
        If strSplitChar = "" Then
            '�޷ָ���ʱ
            strTmp = strInput
            Do While Len(strTmp) > lngLength
                ReDim Preserve strArray(UBound(strArray) + 1)
                strArray(UBound(strArray)) = Mid(strTmp, 1, lngLength)
                strTmp = Mid(strTmp, lngLength + 1)
            Loop
            
            If strTmp <> "" Then
                ReDim Preserve strArray(UBound(strArray) + 1)
                strArray(UBound(strArray)) = strTmp
            End If
        Else
            '�зָ���ʱ
            ArrTmp = Split(strInput & strSplitChar, strSplitChar)
            lngCount = UBound(ArrTmp)
        
            For i = 0 To lngCount
                If ArrTmp(i) <> "" Then
                    '�зָ�������Ҫ���ַָ���֮���ַ��������ԣ����ܰѷָ���֮����ַ���
                    If Len(IIf(strTmp = "", "", strTmp & strSplitChar) & ArrTmp(i)) > lngLength Then
                        ReDim Preserve strArray(UBound(strArray) + 1)
                        strArray(UBound(strArray)) = strTmp
                        strTmp = ArrTmp(i)
                    Else
                        strTmp = IIf(strTmp = "", "", strTmp & strSplitChar) & ArrTmp(i)
                    End If
                End If
                       
                If i = lngCount Then
                    ReDim Preserve strArray(UBound(strArray) + 1)
                    strArray(UBound(strArray)) = strTmp
                End If
            Next
        End If
    Else
        ReDim Preserve strArray(UBound(strArray) + 1)
        strArray(UBound(strArray)) = strInput
    End If
    
    Piva_GetArrayByStr = strArray
End Function
Private Sub Piva_IniTransRec()
    '��Һ����¼��
    Set mrsTrans = New ADODB.Recordset
    With mrsTrans
        If .State = 1 Then .Close
        
        '�ü�¼��Ӧ����Һ��ҩ��¼��Ϣ
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҩid", adDouble, 18, adFldIsNullable
        .Fields.Append "����id", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 3, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�Ա�", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "סԺ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���˲���", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "���˿���", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "ִ��ʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����id", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҳid", adDouble, 18, adFldIsNullable
        .Fields.Append "���ȼ�", adDouble, 18, adFldIsNullable
        .Fields.Append "���˿���id", adDouble, 18, adFldIsNullable
        .Fields.Append "���ʱ��", adLongVarChar, 20, adFldIsNullable
        
        '��Һ��ҩ��¼ҵ�������Ϣ
        .Fields.Append "��ҩ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ƿǩ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ӡ��־", adDouble, 1, adFldIsNullable
        .Fields.Append "�Ƿ���", adDouble, 1, adFldIsNullable
        .Fields.Append "�˲���", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�˲�ʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ҩ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ҩʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ҩ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ҩ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ҩʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��������ʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�������ʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����ҩ��", adLongVarChar, 1, adFldIsNullable
        .Fields.Append "ҩʦ���ʱ��", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "�Ƿ��������", adDouble, 1, adFldIsNullable
        .Fields.Append "�Ƿ�����", adDouble, 1, adFldIsNullable
        .Fields.Append "����ҩ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�ֹ���������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����ԭ��", adLongVarChar, 200, adFldIsNullable
        
        '��Һ��ҩ��¼��Ӧ��ҩƷ��Ϣ
        .Fields.Append "�շ�id", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 2, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "ҩƷ����", adLongVarChar, 50, adFldIsNullable   '����+ͨ����/��Ʒ��
        .Fields.Append "ͨ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��Ʒ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "Ӣ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "������λ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "Ƶ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "��λ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "�÷�", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩ��id", adDouble, 18, adFldIsNullable
        .Fields.Append "�������", adDouble, 3, adFldIsNullable
        .Fields.Append "����id", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҩ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ý", adDouble, 1, adFldIsNullable
        
        .Fields.Append "��ҩ����", adDouble, 18, adFldIsNullable
        .Fields.Append "�������", adDouble, 18, adFldIsNullable
        .Fields.Append "ʵ������", adDouble, 18, adFldIsNullable
        
        .Fields.Append "�����", adDouble, 1, adFldIsNullable
        .Fields.Append "ҽ������ʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ҽ��id", adDouble, 18, adFldIsNullable
        .Fields.Append "���ͺ�", adDouble, 18, adFldIsNullable
        
        .Fields.Append "ִ�б�־", adDouble, 1, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adDouble, 5, adFldIsNullable
        .Fields.Append "��ɫ", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub Piva_IniPriRec()
    Set mrstemp = New ADODB.Recordset
    With mrstemp
        If .State = 1 Then .Close
        .Fields.Append "��ҩid", adDouble, 18, adFldIsNullable
        .Fields.Append "����id", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҩ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "Ƶ��", adLongVarChar, 20, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub Piva_Set���ȼ�(ByVal rstemp As Recordset, ByRef rsTrans As Recordset, ByVal lng��ҩid As Long)
    Dim lng���ȼ�A As Long
    Dim lng���ȼ�B As Long
    Dim lng���ȼ�C As Long
    Dim lng���ȼ�D As Long
    Dim bln�������� As Boolean
    Dim bln����Ƶ�� As Boolean
    
    If rstemp.EOF Or mrsPRI.EOF Then Exit Sub
    mrsPRI.MoveFirst
    mrsPRI.Sort = "���ȼ�"
    
    mrsPRI.Filter = "����id='" & rstemp!����ID & "'"

    rsTrans.Filter = "��ҩid=" & lng��ҩid
    rsTrans.Sort = ""
    
    If mrsPRI.EOF Then mrsPRI.Filter = "����id='0'"
    Do While Not mrsPRI.EOF
        rsTrans.MoveFirst
        Do While Not rsTrans.EOF
            If mrsPRI!��ҩ���� = rsTrans!��ҩ���� Then
                bln�������� = True
                If Mid(mrsPRI!Ƶ��, 1, IIf(InStr(1, mrsPRI!Ƶ��, "(") = 0, 1, InStr(1, mrsPRI!Ƶ��, "(") - 1)) = rsTrans!Ƶ�� Then
                    bln����Ƶ�� = True
                    lng���ȼ�A = mrsPRI!���ȼ�
                ElseIf mrsPRI!Ƶ�� = "����Ƶ��" And Not bln����Ƶ�� Then
                    lng���ȼ�B = mrsPRI!���ȼ�
                ElseIf mrsPRI!Ƶ�� = "����Ƶ��" And Not bln����Ƶ�� Then
                    lng���ȼ�B = mrsPRI!���ȼ�
                End If
            ElseIf mrsPRI!��ҩ���� = "��������" And Not bln�������� Then
                If Mid(mrsPRI!Ƶ��, 1, IIf(InStr(1, mrsPRI!Ƶ��, "(") = 0, 1, InStr(1, mrsPRI!Ƶ��, "(") - 1)) = rsTrans!Ƶ�� Then
                    bln����Ƶ�� = True
                    lng���ȼ�C = mrsPRI!���ȼ�
                ElseIf mrsPRI!Ƶ�� = "����Ƶ��" And Not bln����Ƶ�� Then
                    lng���ȼ�D = mrsPRI!���ȼ�
                ElseIf mrsPRI!Ƶ�� = "����Ƶ��" And Not bln����Ƶ�� Then
                    lng���ȼ�D = mrsPRI!���ȼ�
                End If
            End If
            rsTrans.MoveNext
        Loop
        mrsPRI.MoveNext
    Loop
    
    rsTrans.MoveFirst
    Do While Not rsTrans.EOF
        rsTrans!���ȼ� = IIf(bln��������, IIf(bln����Ƶ��, lng���ȼ�A, lng���ȼ�B), IIf(bln����Ƶ��, lng���ȼ�C, lng���ȼ�D))
        rsTrans.Update
        rsTrans.MoveNext
    Loop
    rsTrans.Filter = ""
End Sub

Private Function Piva_GetTrans(ByVal lngCenterID As Long, ByVal strSendNO As String) As ADODB.Recordset
    'ȡ��Һ��ҩ��¼
    'lngCenterID����Һ��������ID
    'str����ID������ID��
    'dateExeStart��dateExeEnd����Һ��ҩ���ݵ�ִ��ʱ�䷶Χ
    On Error GoTo errHandle
    
    gstrSQL = "Select Distinct A.ID As ��ҩID, A.����id, A.���, A.��ҩ����, S.��ɫ,A.����, A.�Ա�, A.����, A.סԺ��, A.����,M.ҩʦ���ʱ��, A.���˲���id, A.���˿���id, A.ִ��ʱ��, A.ƿǩ��,A.���ʱ��,M.ִ��Ƶ��,A.�Ƿ��������,A.�Ƿ�����,A.�ֹ���������,'' ����ԭ��," & _
        "  A.������Ա,A.����ʱ��,Nvl(A.��ӡ��־,0) As ��ӡ��־, A.�Ƿ���, B.���� As ���˲���, C.���� As ���˿���, D.�շ�id, E.����, E.NO, F.���� As ҩƷ����, " & _
        " F.���� As ͨ����, H.���� As ��Ʒ��, I.���� As Ӣ����, F.���, E.����, E.����, E.����, J.���㵥λ As ������λ,J.id ҩ��id, E.Ƶ��, '' As ��������, " & _
        " Case Nvl(E.�����, 'δ���') When 'δ���' Then E.ʵ������ * Nvl(E.����, 1) / G.סԺ��װ Else 0 End As ��ҩ����,M.����id,M.��ҳid,T.��ý,A.ҽ��id,A.���ͺ�, " & _
        " (D.���� / G.סԺ��װ)  As ����,D.���� As ʵ������, G.סԺ��λ As ��λ,Nvl(E.����,0) As ����, Nvl(L.ʵ������, 0)/ G.סԺ��װ As �������, Nvl(M.�����,-1) �����, E.�÷�, E.ҩƷid, n.��� As �������,E.����id, o.����, A.��ҩ����,r.����ʱ�� As ҽ������ʱ��,nvl(T.������,'0') ��ҩ���� " & _
        " From  ��Һ��ҩ��¼ A, ���ű� B, ���ű� C, ��Һ��ҩ���� D, ҩƷ�շ���¼ E, �շ���ĿĿ¼ F, ҩƷ��� G,��ҺҩƷ���� X,  �շ���Ŀ���� H, ������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M, סԺ���ü�¼ N, ������ҳ O ,��ҩ�������� S,ҩƷ���� T " & _
        ",(Select �ⷿid, ҩƷid, Nvl(����, 0) As ����, Nvl(ʵ������, 0) As ʵ������ " & _
        " From ҩƷ��� Where ���� = 1 And �ⷿid = [1]) L, ҩƷ�շ���¼ P, ����ҽ������ R " & _
        " Where A.���˲���id = B.ID And A.���˿���id = C.ID And A.ID = D.��¼id And D.�շ�id = E.ID And E.ҩƷid = F.ID And F.ID = G.ҩƷid And G.ҩƷid=X.ҩƷid(+) And E.����id = N.ID And N.ҽ����� = M.ID And " & _
        " G.ҩƷid = H.�շ�ϸĿid(+) And H.����(+) = 3 And G.ҩ��id = I.������Ŀid(+) And I.����(+) = 2 And G.ҩ��id = J.ID And T.ҩ��id=J.ID And A.��ҩ����=S.����(+) And E.�ⷿid = L.�ⷿid(+) And E.ҩƷid = L.ҩƷid(+) And Nvl(E.����, 0) = L.����(+) " & _
        " And n.����id = o.����id(+) And n.��ҳid = o.��ҳid(+) And A.����id = [1] And a.ҽ��id = r.ҽ��id And a.���ͺ� = r.���ͺ� " & _
        " And e.���� = p.���� And e.No = p.No And e.�ⷿid + 0 = p.�ⷿid And e.ҩƷid + 0 = p.ҩƷid+0 And e.��� = p.��� And (p.��¼״̬ = 1 Or Mod(p.��¼״̬, 3) = 0) " & _
        " And A.����״̬=1 And R.���ͺ� = [2] "
     
    Set Piva_GetTrans = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Һ��ҩ��¼", lngCenterID, strSendNO)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub Piva_Set����(ByRef rstemp As Recordset, ByVal lng����ID As Long, ByVal str���� As String, ByVal strִ��ʱ�� As String)
    Dim lng���� As Long
    Dim lng���� As Long
    Dim lngRow As Long
    Dim bln���� As Boolean
    Dim lng��ҩid As Long
    Dim rs���� As Recordset
    Dim str��ҩ���� As String
    Dim strCon As String
    Dim lngOld��ҩid As Long
    Dim blnLoop As Boolean
    Dim lng���� As Long
    Dim strOld���� As String
    Dim strC���� As String
    
    If rstemp.EOF Then Exit Sub
    
    rstemp.MoveFirst
    rstemp.Sort = "����ҩ����,���ȼ�,ִ��ʱ��,��ҩid"
    
    rstemp.Filter = "����id=" & lng����ID & " "
    rstemp.MoveFirst
    Do While Not rstemp.EOF
        blnLoop = False
        If Format(rstemp!ִ��ʱ��, "YYYY-MM-DD") = strִ��ʱ�� Then
            If rstemp!����ҩ���� = strOld���� Then
                If rstemp!��ý = 1 Then lng���� = lng���� + rstemp!����
            Else
                strOld���� = rstemp!����ҩ����
                lng���� = 0
                mrsVol.MoveFirst
                Do While Not mrsVol.EOF
                    If (mrsVol!����ID = "0" Or Val(mrsVol!����ID) = rstemp!���˿���id Or mrsVol!����ID = "00") And (mrsVol!��ҩ���� = rstemp!����ҩ���� Or mrsVol!��ҩ���� = "") Then
                        If Val(mrsVol!����) > lng���� Then
                            lng���� = Val(mrsVol!����)
                        End If
                    End If
                    mrsVol.MoveNext
                Loop
                
                lng���� = 0
                lng��ҩid = 0
                If rstemp!��ý = 1 Then lng���� = rstemp!����
            End If
            
            
            If lng��ҩid <> rstemp!��ҩid Then
                lng���� = 0
                strCon = strCon & " And ��ҩid<>" & rstemp!��ҩid
                If rstemp!��ý = 1 Then lng���� = rstemp!����
                lng��ҩid = rstemp!��ҩid
                lngRow = lngRow + 1
            Else
                If rstemp!��ý = 1 Then lng���� = lng���� + rstemp!����
            End If
        
            If lngRow > 1 And lng���� > 0 Then
                If lng���� - lng���� > lng���� And lng���� <> 0 Then
                    lng���� = lng���� - lng����
                    rstemp.Filter = "��ҩid=" & lng��ҩid & ""
                    rstemp.Sort = ""
                     
                    rstemp.MoveFirst
                    Do While Not rstemp.EOF
                        rstemp!����ҩ���� = Val(Mid(rstemp!����ҩ����, 1, Len(rstemp!����ҩ����) - 1)) + 1 & "#"
                        rstemp.Update
                        rstemp.MoveNext
                    Loop
                    
                    strC���� = strOld����
                    If strCon <> "" Then strCon = Mid(strCon, 1, Len(strCon) - Len(" And ��ҩid<>" & lng��ҩid))
                    
                    rstemp.Filter = "����id=" & lng����ID & strCon
                    rstemp.Sort = "����ҩ����,���ȼ�,ִ��ʱ��,��ҩid"
                    If rstemp.RecordCount <> 0 Then rstemp.MoveFirst
                    blnLoop = True
                End If
            End If
        Else
            strCon = strCon & " And ��ҩid<>" & rstemp!��ҩid
        End If
        If Not rstemp.EOF And Not blnLoop Then rstemp.MoveNext
    Loop
    rstemp.Filter = ""
End Sub



Private Sub Piva_SetLastBatch(ByRef rsTrans As ADODB.Recordset)
    '�����ϴ����Σ���ҽ���仯(�¿�,��ͣ)ʱ������
    Dim lngҽ��id As Long
    Dim lng���ͺ� As Long
    Dim strִ��ʱ�� As String
    Dim rsData As ADODB.Recordset
    Dim strOldFilter, strOldSort As String
    Dim str����ids As String
    Dim lng����ID As Long
    Dim str�������β���ids As String
    Dim str����ҽ��ids As String
    Dim str�ϴ�ҽ��ids As String
    Dim intCount As Integer
    Dim i As Integer
    
    '��¼��ʼ�Ĺ��˼���������
    strOldFilter = rsTrans.Filter
    strOldSort = rsTrans.Sort
    
    'ͳ�Ʋ���
    With rsTrans
        .Filter = ""
        If .RecordCount = 0 Then Exit Sub
        .Sort = "����id"
        
        Do While Not .EOF
            If lng����ID <> !����ID Then
                lng����ID = !����ID
            
                gstrSQL = "Select a.ҽ��id, a.���ͺ� From ��Һ��ҩ��¼ A, ����ҽ����¼ B,������ĿĿ¼ C Where a.ҽ��id = b.Id and B.������Ŀid=C.id and c.��������=2 and b.�������='E' And b.����id = [1] and b.��ҳid =[2]"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "SetLastBatch", lng����ID, !��ҳid)
                
                str����ҽ��ids = ""
                str�ϴ�ҽ��ids = ""
                lng���ͺ� = 0
                intCount = 0
                With rsData
                    .Sort = "���ͺ� Desc,ҽ��id"
                    
                    Do While Not .EOF
                        If lng���ͺ� <> !���ͺ� Then
                            intCount = intCount + 1
                            lng���ͺ� = !���ͺ�
                        End If
                        
                        If intCount = 1 Then
                            'ȡ���η��͵�ҽ��ID
                            If InStr(1, str����ҽ��ids, !ҽ��id) = 0 Then
                                str����ҽ��ids = IIf(str����ҽ��ids = "", "", str����ҽ��ids & ",") & !ҽ��id
                            End If
                        ElseIf intCount = 2 Then
                            'ȡ�ϴη��͵�ҽ��ID
                            If InStr(1, str�ϴ�ҽ��ids, !ҽ��id) = 0 Then
                                str�ϴ�ҽ��ids = IIf(str�ϴ�ҽ��ids = "", "", str�ϴ�ҽ��ids & ",") & !ҽ��id
                            End If
                        Else
                            Exit Do
                        End If
                        
                        rsData.MoveNext
                    Loop
                End With
            
                '���η���ҽ����ҽ��IDһ������ʾû�б仯
                If str����ҽ��ids = str�ϴ�ҽ��ids Then
                    str�������β���ids = IIf(str�������β���ids = "", "", str�������β���ids & ",") & lng����ID
                End If
            End If
        .MoveNext
        Loop
    End With
    
    If str�������β���ids = "" Then Exit Sub
    
    '�����ϴ���������
    For i = 0 To UBound(Split(str�������β���ids, ","))
         With rsTrans
            .Filter = "����id=" & Split(str�������β���ids, ",")(i)
            .Sort = "��ҩID"
            
            Do While Not .EOF
                lngҽ��id = !ҽ��id
                lng���ͺ� = !���ͺ�
                strִ��ʱ�� = IIf(IsNull(!ִ��ʱ��), "", Format(!ִ��ʱ��, "YYYY-MM-DD HH:MM:SS"))
    
                gstrSQL = " Select Distinct ��ҩ���� " & _
                    " From ��Һ��ҩ��¼ A " & _
                    " Where ҽ��id = [1] And ���ͺ� = (Select Distinct Max(���ͺ�) From ��Һ��ҩ��¼ Where ҽ��id = [1] And ���ͺ� <> [2]) And " & _
                    " To_Char(a.ִ��ʱ��, 'hh24:mi:ss') = To_Char([3], 'hh24:mi:ss') "
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "SetLastBatch", lngҽ��id, lng���ͺ�, CDate(strִ��ʱ��))
                
                If rsData.RecordCount > 0 Then
                    !����ҩ���� = rsData!��ҩ���� & "#"
                    .Update
                End If
                
                .MoveNext
            Loop
        End With
    Next
    
    '�ָ���¼��״̬
    rsTrans.Filter = IIf(strOldFilter = "0", 0, strOldFilter)
    rsTrans.Sort = strOldSort
End Sub



