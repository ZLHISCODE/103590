Attribute VB_Name = "mdlDIH"
Option Explicit

Private Function GetXML_RecipeDetail_DIH(ByVal LngStockID As Long, ByVal strNO As String) As String
'��������ϸ��֯��ָ����XML��ʽ
'���ýӿڣ���������(DIH)
    Dim rsRecipe As Recordset   '���˺ʹ�����¼
    Dim rsDiagnosis As Recordset    '��ϼ�¼
    Dim rsDrug As Recordset         '����ҩƷ��¼
    Dim strSql As String
    Dim strRecipe As String
    Dim strDiagnosis As String
    Dim strXML As String
    Dim strXML_Patient As String
    Dim strXML_Recipe As String
    Dim strXML_Drug As String
    Dim i As Integer
    Dim strOutput As String
    Dim strOutPutExeStep As String    'ִ�в��裬���������־�����������
    Dim strTmp As String
    
    strOutput = strOutput & vbCrLf & "����GetXML_RecipeDetail_DIH"
    
    On Error GoTo errHandle
    
    '�жϵ��������Ƕദ��
    If InStr(1, strNO, "|") < 1 Then
        '������
        strRecipe = " And a.����=[2] And a.NO=[3] "
    Else
        '�ദ��
        strRecipe = " And ("
        For i = 0 To UBound(Split(strNO, "|"))
            If i = UBound(Split(strNO, "|")) Then
                strRecipe = strRecipe & "(a.����=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "')"
            Else
                strRecipe = strRecipe & "(a.����=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "') or "
            End If
        Next
        strRecipe = strRecipe & ") "
    End If
    
    strOutPutExeStep = "�жϺͷֽⵥ����/�ദ�����"
    
    '��ȡ���˼���������Ϣ
    strSql = "Select Distinct a.����id, decode(c.����,null,d.����,c.����) ����, decode(c.�Ա�,null,d.�Ա�,c.�Ա�) �Ա�, decode(c.����,null,d.����,c.����) ����, " & vbNewLine & _
        "       c.���, c.ҽ�Ƹ��ʽ ҽ������, d.�ѱ� As �շ����, a.No As ������, Decode(a.��������, 2, 'J', 'M') As ��������, d.NO As ������, " & vbNewLine & _
        "       d.��������id As ������ұ���, f.���� As �����������, g.Id As ����ҽ������, d.������ As ����ҽ������, d.�Ǽ�ʱ�� As �ɷ�ʱ��,a.���� " & vbNewLine & _
        " From δ��ҩƷ��¼ A, ������Ϣ C, ������ü�¼ D, ҩƷ�շ���¼ E, ���ű� F, ��Ա�� G " & vbNewLine & _
        " Where a.���� = e.���� And a.No = e.No And a.�ⷿid = e.�ⷿid And a.����id = c.����id(+) And e.����id = d.Id And d.��������id = f.Id And " & vbNewLine & _
        "      d.������ = g.���� And a.�ⷿid = [1] " & strRecipe
    
    strOutPutExeStep = strOutPutExeStep & vbCrLf & "��ѯ���˼���������Ϣ��ʼ��" & vbCrLf & strSql
    
    If gintMode = 0 Then
        Set rsRecipe = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "GetXML_RecipeDetail_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    Else
        Set rsRecipe = mdlDrugPacker.OpenSQLRecord(strSql, "GetXML_RecipeDetail_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    End If
    
'�����ʱ���ϴ�������ע��
'    '��ȡ�����Ϣ
'    strSql = "Select d.�������, d.�Ƿ�����, a.����, a.No" & vbNewLine & _
'        " From δ��ҩƷ��¼ A, ������ü�¼ B, ҩƷ�շ���¼ C, ������ϼ�¼ D, �������ҽ�� E" & vbNewLine & _
'        " Where a.���� = c.���� And a.No = c.No And a.�ⷿid = c.�ⷿid And c.����id = b.Id And e.ҽ��id = b.ҽ����� And d.Id = e.���id And" & vbNewLine & _
'        " d.ȡ��ʱ�� Is Null And a.�ⷿid = [1] " & strRecipe
'
'    strOutPutExeStep = strOutPutExeStep & vbCrLf & "��ѯ�����Ϣ��ʼ��" & vbCrLf & strSql
'
'    If gintMode = 0 Then
'        Set rsDiagnosis = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "GetXML_RecipeDetail_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
'    Else
'        Set rsDiagnosis = mdlDrugPacker.OpenSQLRecord(strSql, "GetXML_RecipeDetail_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
'    End If
    
    '��ȡ����ҩƷ��Ϣ
    strSql = "Select Distinct a.����, a.No, b.Id ҩƷ����, b.���� ҩƷ����, b.��� ҩƷ���, a.���� ҩƷ����, a.ʵ������ / d.�����װ ҩƷ����, d.���ﵥλ ��ҩ��λ, a.�÷� As ���÷���," & vbNewLine & _
        " a.����, g.���㵥λ, f.ִ��Ƶ��, f.ҽ������ as ��ע˵��, a.�ⷿid As ҩ������, a.���" & vbNewLine & _
        " From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B, ҩƷ��� D, ������ü�¼ E, ����ҽ����¼ F, ������ĿĿ¼ G" & vbNewLine & _
        " Where a.ҩƷid = b.Id And a.ҩƷid = d.ҩƷid And a.����id = e.Id And d.ҩ��id = g.Id And e.ҽ����� = f.Id(+) And a.�ⷿid = [1] " & strRecipe
    
    strOutPutExeStep = strOutPutExeStep & vbCrLf & "��ѯ����ҩƷ��Ϣ��ʼ��" & vbCrLf & strSql
    
    If gintMode = 0 Then
        Set rsDrug = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "GetXML_RecipeDetail_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    Else
        Set rsDrug = mdlDrugPacker.OpenSQLRecord(strSql, "GetXML_RecipeDetail_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    End If
    
    strOutput = strOutput & vbCrLf & "��ѯ������ϸ���"
    
'    1.1.1�����ﴦ��xml��ʽ:
'    <outpOrder>
'        <patient>                            - ������Ϣ
'        <windowNo></windowNo>   - ȡҩ���ںţ�ֻ���� ����HIS ���䴰�ں�ʱ�У��Ҳ���Ϊ 0��
'        <patientID></patientID>           - ����Ψһ ID
'        <patientName></patientName>       - ����
'        <patientGender></patientGender>   -�Ա�
'        <patientAge></patientAge>          -����
'        <identity></identity>                 -���
'        <insuranceType></insuranceType>    -ҽ������
'        <chargeType></chargeType>         -�շ����
'        </patient>
'        <prescriptions>  - �����嵥
'            <prescription no="" type="" paymentDT="">   - ������no������Ψһ��ţ�type��M-���J-���O-�������ɷ�ʱ�䣺yyyy-MM-dd HH:mm:ss
'            <outpNo></outpNo>                 - ������
'            <visitNo></visitNo>               - ������
'            <deptCode></deptCode>             - ������ұ���
'            <deptName></deptName>             - �����������
'            <doctCode></doctCode>             - ����ҽ������
'            <doctName></doctName>             - ����ҽ������
'            <diagnosis></diagnosis>           - �ٴ����
'            <paymentDT></paymentDT>       -�ɷ�ʱ�䣺yyyy-MM-dd HH:mm:ss
'            <drugList>   -������ҩƷ�嵥
'                <drug>
'                <drugCode></drugCode>           -   ҩƷ����
'                <drugName></drugName>       -   ����
'                <drugSpec></drugSpec>          -    ���
'                <firmName></firmName>          -    ����
'                <amount></amount>             - ҩƷ����
'                <takeUnit></takeUnit>             - ��ҩ��λ
'                <takeMethod></takeMethod>      -    ���÷���
'                <takeDosage></takeDosage>       - ����
'                <takeType></takeType>           -   ��������
'                <takeNote></takeNote>          -    ��ע˵��
'                <pharmacyCode></pharmacyCode>   -   ҩ������
'                <sortNo></sortNo>     - �ڴ����е�ҩƷ˳��ţ�����ֵ��
'                </drug>
'            </drugList>
'            </prescription>
'        </prescriptions>
'    </outpOrder>

    Call OutputLog(vbCrLf & strOutPutExeStep & vbCrLf & _
                    "��ز�����" & "LngStockID=" & LngStockID & " strNO=" & strNO)

    If rsRecipe.RecordCount > 0 Then
        rsRecipe.MoveFirst
    
        '������Ϣ
        With rsRecipe
            strOutPutExeStep = "��֯������ϢXML"
            
            strXML_Patient = "<patient>"
            strXML_Patient = strXML_Patient & vbCrLf & GetXMLFormat("windowNo", "", False)
            strXML_Patient = strXML_Patient & vbCrLf & GetXMLFormat("patientID", NVL(!����id), False)
            strXML_Patient = strXML_Patient & vbCrLf & GetXMLFormat("patientName", NVL(!����), False)
            strXML_Patient = strXML_Patient & vbCrLf & GetXMLFormat("patientGender", NVL(!�Ա�), False)
            strXML_Patient = strXML_Patient & vbCrLf & GetXMLFormat("patientAge", NVL(!����), False)
            strXML_Patient = strXML_Patient & vbCrLf & GetXMLFormat("identity", NVL(!���), False)
            strXML_Patient = strXML_Patient & vbCrLf & GetXMLFormat("insuranceType", NVL(!ҽ������), False)
            strXML_Patient = strXML_Patient & vbCrLf & GetXMLFormat("chargeType", NVL(!�շ����), False)
            strXML_Patient = strXML_Patient & "</patient>"
        End With
        
        '������Ϣ
        With rsRecipe
            strXML_Recipe = "<prescriptions>"
            Do While Not .EOF
                strOutPutExeStep = "��֯������ϢXML"
                
                strXML_Recipe = strXML_Recipe & vbCrLf & "<prescription no=""" & NVL(!������) & """ type=""" & NVL(!��������) & """ paymentDT=""" & Format(NVL(!�ɷ�ʱ��), "yyyy-MM-DD hh:mm:ss") & """>"
                strXML_Recipe = strXML_Recipe & vbCrLf & GetXMLFormat("outpNo", "", False)
                strXML_Recipe = strXML_Recipe & vbCrLf & GetXMLFormat("visitNo", NVL(!������), False)
                strXML_Recipe = strXML_Recipe & vbCrLf & GetXMLFormat("deptCode", NVL(!������ұ���), False)
                strXML_Recipe = strXML_Recipe & vbCrLf & GetXMLFormat("deptName", NVL(!�����������), False)
                strXML_Recipe = strXML_Recipe & vbCrLf & GetXMLFormat("doctCode", NVL(!����ҽ������), False)
                strXML_Recipe = strXML_Recipe & vbCrLf & GetXMLFormat("doctName", NVL(!����ҽ������), False)
                strXML_Recipe = strXML_Recipe & vbCrLf & GetXMLFormat("diagnosis", "", False)   '���
                strXML_Recipe = strXML_Recipe & vbCrLf & GetXMLFormat("paymentDT", Format(NVL(!�ɷ�ʱ��), "yyyy-MM-DD hh:mm:ss"), False)
                
                'ҩƷ��Ϣ
                strXML_Drug = "<drugList>"
                rsDrug.Filter = "no='" & !������ & "' and ����=" & NVL(!����)
                rsDrug.Sort = "���"
                
                Do While Not rsDrug.EOF
                    strOutPutExeStep = "��֯ҩƷ��ϢXML"
                    
                    strXML_Drug = strXML_Drug & vbCrLf & "<drug>"
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("drugCode", NVL(rsDrug!ҩƷ����), False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("drugName", NVL(rsDrug!ҩƷ����), False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("drugSpec", NVL(rsDrug!ҩƷ���), False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("firmName", NVL(rsDrug!ҩƷ����), False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("amount", NVL(rsDrug!ҩƷ����), False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("takeUnit", NVL(rsDrug!��ҩ��λ), False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("takeMethod", NVL(rsDrug!���÷���), False)
                    If NVL(rsDrug!����) = "" Then
                        strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("takeDosage", "", False)
                    Else
                        strTmp = Format(rsDrug!����, "#0.##########") & NVL(rsDrug!���㵥λ) & "��" & NVL(rsDrug!ִ��Ƶ��)
                        strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("takeDosage", strTmp, False)
                    End If
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("takeType", "", False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("takeNote", NVL(rsDrug!��ע˵��), False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("pharmacyCode", NVL(rsDrug!ҩ������), False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("sortNo", NVL(rsDrug!���), False)
                    strXML_Drug = strXML_Drug & vbCrLf & "</drug>"
                    
                    rsDrug.MoveNext
                Loop
                
                strXML_Drug = strXML_Drug & vbCrLf & "</drugList>"
                strXML_Recipe = strXML_Recipe & vbCrLf & strXML_Drug & vbCrLf & "</prescription>"
                
                rsRecipe.MoveNext
            Loop
            
            '���ܴ���ҩƷ
            strXML_Recipe = strXML_Recipe & "</prescriptions>"
        End With
        
        '���ܲ��ˡ�������ҩƷ��Ϣ��ƴ��������XML
        strXML = "<outpOrder>"
        strXML = strXML & vbCrLf & strXML_Patient
        strXML = strXML & vbCrLf & strXML_Recipe
        strXML = strXML & vbCrLf & "</outpOrder>"
    Else
        strOutput = strOutput & vbCrLf & "�޴�������"
    End If
    
    GetXML_RecipeDetail_DIH = strXML
    
    strOutput = strOutput & vbCrLf & "��֯������ϢXML��ɣ�" & vbCrLf & strXML
    strOutput = strOutput & vbCrLf & "ִ�гɹ���GetXML_RecipeDetail_DIH"
    Call OutputLog(strOutput)
    
    Exit Function
errHandle:
    strOutput = strOutput & vbCrLf & "�����쳣����:" & Err.Description
    
    If gobjComLib.ErrCenter = 1 Then Resume
    Call gobjComLib.SaveErrLog

    strOutput = strOutput & vbCrLf & "����裺" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "��ز�����" & "LngStockID=" & LngStockID & " strNO=" & strNO
    strOutput = strOutput & vbCrLf & "���SQL" & vbCrLf & strSql
    strOutput = strOutput & vbCrLf & "ִ��ʧ�ܣ�GetXML_RecipeDetail_DIH"
    Call OutputLog(strOutput)
End Function

Private Function GetXML_RecipeReady_DIH(ByVal LngStockID As Long, ByVal strNO As String) As Variant
'����������֯��ָ����XML��ʽ
'ȡҩ֪ͨ/׼����ҩ
'���ýӿڣ���������(DIH)
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim strSql As String
    Dim strOutput As String
    Dim strOutPutExeStep As String    'ִ�в��裬���������־�����������
    
    On Error GoTo errHandle
    
    strOutput = strOutput & vbCrLf & "���ú�����GetXML_RecipeReady_DIH"
   
    strSql = "Select b.���� As ���ں�, a.����id, a.Groupno, a.Ordertype " & _
        " From δ��ҩƷ��¼ A, ��ҩ���� B " & _
        " Where a.��ҩ���� = b.����(+) And a.�ⷿid=[1] "
    
    '�жϵ��������Ƕദ��
    If InStr(1, strNO, "|") < 1 Then
        '������
        strSql = strSql & " And a.����=[2] And a.NO=[3] "
    Else
        '�ദ��
        strSql = strSql & " And ("
        For i = 0 To UBound(Split(strNO, "|"))
            If i = UBound(Split(strNO, "|")) Then
                strSql = strSql & "(a.����=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "')"
            Else
                strSql = strSql & "(a.����=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "') or "
            End If
        Next
        strSql = strSql & ") "
    End If
    
    strOutPutExeStep = "�жϺͷֽⵥ����/�ദ�����"
    
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "GetXML_RecipeReady_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(strSql, "GetXML_RecipeReady_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    End If
    
    strOutPutExeStep = "ִ��SQL���"
    
    With rsTemp
        '���GroupNO�Ƿ���д
        If NVL(!Groupno) = "" Then
            strOutput = strOutput & vbCrLf & "GroupNOδ��д�����豸δ��ҩ��ɡ�"
            Call OutputLog(strOutput)
            Exit Function
        End If
    
'    <outpOrderTake>
'        <windowNo>2</windowNo>              --���ں�
'        <patientID>1042323</patientID>      --����ΨһID
'        <groupNo>M15091800973</groupNo>     --���
'        <orderType>indirect</orderType>     --����ֱ������Ԥ�䷢��ʶ
'    </outpOrderTake>

        If .RecordCount > 0 Then
            strXML = "<outpOrderTake>"
            strXML = strXML & vbCrLf & GetXMLFormat("windowNo", NVL(!���ں�), False)
            strXML = strXML & vbCrLf & GetXMLFormat("patientID", NVL(!����id), False)
            strXML = strXML & vbCrLf & GetXMLFormat("groupNo", NVL(!Groupno), False)
            strXML = strXML & vbCrLf & GetXMLFormat("orderType", NVL(!orderType), False)
            strXML = strXML & vbCrLf & "</outpOrderTake>"
        End If
        
        strOutPutExeStep = "ƴװXML���"
    End With
    
    GetXML_RecipeReady_DIH = strXML
    
    strOutput = strOutput & vbCrLf & "��֯XML��ɣ�" & vbCrLf & strXML
    strOutput = strOutput & vbCrLf & "ִ�гɹ���GetXML_RecipeReady_DIH"
    Call OutputLog(strOutput)
    
    Exit Function
    
errHandle:
    strOutput = strOutput & vbCrLf & "�����쳣����" & Err.Description
    
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
    
    strOutput = strOutput & vbCrLf & "����裺" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "��ز�����" & "LngStockID=" & LngStockID & " strNO=" & strNO
    strOutput = strOutput & vbCrLf & "���SQL" & vbCrLf & strSql
    strOutput = strOutput & vbCrLf & "ִ��ʧ�ܣ�GetXML_RecipeReady_DIH"
    Call OutputLog(strOutput)
End Function

Private Function GetXML_RecipeCompletion_DIH(ByVal LngStockID As Long, ByVal strNO As String) As String
'����������֯��ָ����XML��ʽ
'��ҩ���
'���ýӿڣ���������(DIH)
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim strSql As String
    Dim strOutput As String
    Dim strOutPutExeStep As String    'ִ�в��裬���������־�����������
    
    On Error GoTo errHandle
    
    strOutput = strOutput & vbCrLf & "���ú�����GetXML_RecipeCompletion_DIH"
   
    strSql = "Select a.����id, a.Groupno " & _
        " From δ��ҩƷ��¼ A " & _
        " Where a.�ⷿid=[1] "
    
    '�жϵ��������Ƕദ��
    If InStr(1, strNO, "|") < 1 Then
        '������
        strSql = strSql & " And a.����=[2] And a.NO=[3] "
    Else
        '�ദ��
        strSql = strSql & " And ("
        For i = 0 To UBound(Split(strNO, "|"))
            If i = UBound(Split(strNO, "|")) Then
                strSql = strSql & "(a.����=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "')"
            Else
                strSql = strSql & "(a.����=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "') or "
            End If
        Next
        strSql = strSql & ") "
    End If
    
    strOutPutExeStep = "�жϺͷֽⵥ����/�ദ�����"
    
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "GetXML_RecipeCompletion_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(strSql, "GetXML_RecipeCompletion_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    End If
    
    strOutPutExeStep = "ִ��SQL���"
    
    With rsTemp
'        <outpOrderCompletion>
'            <patientID>103278</patientID>      --����ΨһID
'            <groupNo>M15092201309</groupNo>     --���
'        </outpOrderCompletion>

        If .RecordCount > 0 Then
            strXML = "<outpOrderCompletion>"
            strXML = strXML & vbCrLf & GetXMLFormat("patientID", NVL(!����id), False)
            strXML = strXML & vbCrLf & GetXMLFormat("groupNo", NVL(!Groupno), False)
            strXML = strXML & vbCrLf & "</outpOrderCompletion>"
        End If
        
        strOutPutExeStep = "ƴװXML���"
    End With
    
    GetXML_RecipeCompletion_DIH = strXML
    
    strOutput = strOutput & vbCrLf & "��֯XML��ɣ�" & vbCrLf & strXML
    strOutput = strOutput & vbCrLf & "ִ�гɹ���GetXML_RecipeCompletion_DIH"
    Call OutputLog(strOutput)
    
    Exit Function
errHandle:
    strOutput = strOutput & vbCrLf & "�����쳣����" & Err.Description
    
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
    
    strOutput = strOutput & vbCrLf & "����裺" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "��ز�����" & "LngStockID=" & LngStockID & " strNO=" & strNO
    strOutput = strOutput & vbCrLf & "���SQL" & vbCrLf & strSql
    strOutput = strOutput & vbCrLf & "ִ��ʧ�ܣ�GetXML_RecipeCompletion_DIH"
    Call OutputLog(strOutput)
End Function

Public Function HisTransData_DIH(ByVal lngOper As Long, ByVal LngStockID As Long, ByVal strNO As String, ByRef strReturn As String) As Boolean
    '�ϴ�HIS���ݵ��Է�ϵͳ�����ϴ��ؼ���Ϣ�����նԷ�������Ϣ
    Dim strInputXML As String
    Dim strOutput As String
    Dim strTmp As String
    Dim strOutXML As String
    Dim strOut_RETCODE As String
    Dim strOut_WinNo As String
    Dim strOut_Msg As String
    Dim strOutPutExeStep As String
    Dim objXML As clsXML
    
    strOutput = strOutput & vbCrLf & "���ú�����HisTransData_DIH"
    strOutput = strOutput & vbCrLf & "lngOper=" & lngOper
    strOutput = strOutput & vbCrLf & "strNO=" & strNO
    
    On Error GoTo errHandle
    
    'ҵ����룺
    Select Case lngOper
        Case gType.IntDetail
            '�ϴ�������ϸ
            strInputXML = GetXML_RecipeDetail_DIH(LngStockID, strNO)
            
            If strInputXML = "" Then Exit Function
            
            '���öԷ��ӿ�
            strOutPutExeStep = "���öԷ��ӿڿ�ʼ��outpOrderDispense"
            strOutXML = gobjSOAP.outpOrderDispense(strInputXML)
            strOutPutExeStep = "���öԷ��ӿ���ɣ�outpOrderDispense"
            
            strOutput = strOutput & vbCrLf & "���öԷ��ӿ���ɣ�outpOrderDispense"
        Case gType.IntStartList
            'ȡҩ֪ͨ/׼����ҩ
            strInputXML = GetXML_RecipeReady_DIH(LngStockID, strNO)
            
            If strInputXML = "" Then Exit Function
            
            '���öԷ��ӿ�
            strOutPutExeStep = "���öԷ��ӿڿ�ʼ��outpOrderTakeNotify"
            strOutXML = gobjSOAP.outpOrderTakeNotify(strInputXML)
            strOutPutExeStep = "���öԷ��ӿ���ɣ�outpOrderTakeNotify"
            
            strOutput = strOutput & vbCrLf & "���öԷ��ӿ���ɣ�outpOrderTakeNotify"
        Case gType.IntEndList
            HisTransData_DIH = True
            Exit Function
            
            '��������Ҫ�����ε��÷�ҩ�ɹ��ӿڣ���Ϊ�����и���ѭ������������¡�
        
            '��ҩ���
            strInputXML = GetXML_RecipeCompletion_DIH(LngStockID, strNO)
            
            If strInputXML = "" Then Exit Function
            
            '���öԷ��ӿ�
            strOutPutExeStep = "���öԷ��ӿڿ�ʼ��outpOrderCompletionNotify"
            strOutXML = gobjSOAP.outpOrderCompletionNotify(strInputXML)
            strOutPutExeStep = "���öԷ��ӿ���ɣ�outpOrderCompletionNotify"
            
            strOutput = strOutput & vbCrLf & "���öԷ��ӿ���ɣ�outpOrderCompletionNotify"
    End Select
        
    strOutput = strOutput & vbCrLf & "������Ϣ��" & vbCrLf & strOutXML
                            
    '������Ϣ��ʽ
    ''�ϴ�������ϸ
    '    <result>
    '    <status code="0" message="OK"/> -  code������Ϊ�����ţ�message:     �������
    '    <value> - �ӿڷ��ؽ������
    '    <windowNo>2</windowNo> - DIHϵͳ�����ȡҩ���ں�,�������ʧ�ܷ���Ĭ�ϵ�һ�����ں�
    '    </value>
    '    </result>
    
    ''ȡҩ֪ͨ/׼����ҩ
    '<result>
    '<status code="0" message=""/>       - code������Ϊ�����ţ�message���������
    '</result>
    ''��ҩ���
    '�����ַ�0���0
    
    '�ϴ�������ϸ��ȡҩ֪ͨ/׼����ҩ �Ż᷵��xml�ṹ�ַ���
    If lngOper = gType.IntDetail Or lngOper = gType.IntStartList Then
        '������������
        Set objXML = New clsXML
        If objXML.OpenXMLDocument(strOutXML) = False Then
            strOut_Msg = "HisTransData_DIH��������MSXML2.DOMDocument��ʧ�ܣ�"
            If gblnShowMsg Then
                MsgBox strOut_Msg, vbInformation + vbOKOnly, GSTR_MESSAGE
            Else
                strReturn = strOut_Msg
            End If
            
            Call OutputLog(strOutput & vbNewLine & strOut_Msg & vbNewLine)
            Exit Function
        End If
        
        '��ȡcodeֵ
        strOut_RETCODE = objXML.GetXMLNodePropertyValue("status", "code")
        
        '��ȡmessageֵ
        strOut_Msg = objXML.GetXMLNodePropertyValue("status", "message")
        
        If lngOper = gType.IntDetail Then
            '��ȡ���ط�ҩ���ں�
            strOut_WinNo = objXML.GetXMLNodePropertyValue("result/value", "windowNo")
        End If
        
        '�ͷ�XML����
        If Not objXML Is Nothing Then
            objXML.CloseXMLDocument
            Set objXML = Nothing
        End If
    
        strOutPutExeStep = "�������ز������"
    Else
        strOut_RETCODE = Trim(strOutXML)
    End If
           
    '����0��ʾ�ӿڵ��óɹ�������ֵΪ���ɹ�
    If strOut_RETCODE <> "0" Then
        If gblnShowMsg Then
            MsgBox strOut_Msg, vbInformation + vbOKOnly, GSTR_MESSAGE
        Else
            strReturn = strOut_Msg
        End If
        
        strOutput = strOutput & vbCrLf & "�ϴ�����ʧ�ܣ�" & vbCrLf & strInputXML
        strOutput = strOutput & vbCrLf & "������Ϣ��" & strOut_Msg
        strOutput = strOutput & vbCrLf & "ִ��ʧ�ܣ�HisTransData_DIH"
        Call OutputLog(strOutput)

        Exit Function
    End If

    strOutput = strOutput & vbCrLf & "����code��" & strOut_RETCODE
    strOutput = strOutput & vbCrLf & "����message��" & strOut_Msg
    
    If lngOper = gType.IntDetail Then
        strOutput = strOutput & vbCrLf & "����windowno��" & strOut_WinNo
        If Not SetSendWin(LngStockID, strNO, Val(strOut_WinNo)) Then
            If gblnShowMsg Then
                MsgBox "���������ķ�ҩ����ʧ�ܣ�", vbCritical, GSTR_MESSAGE
            Else
                strReturn = "���������ķ�ҩ����ʧ�ܣ�"
            End If

            strOutput = strOutput & vbCrLf & "���������ķ�ҩ����ʧ�ܣ�"
            Call OutputLog(strOutput)
            Exit Function
        End If
    End If
            
    HisTransData_DIH = True
        
    strOutput = strOutput & vbCrLf & "ִ�гɹ���HisTransData_DIH"
    Call OutputLog(strOutput)
    
    Exit Function
    
errHandle:
    strOutput = strOutput & vbCrLf & "�����쳣����" & Err.Description
    
    If gblnShowMsg Then
        If gintMode = 0 Then
            If gobjComLib.ErrCenter = 1 Then Resume
            Call gobjComLib.SaveErrLog
        Else
            MsgBox Err.Description, vbInformation, GSTR_SYSNAME
        End If
    End If
    
    strOutput = strOutput & vbCrLf & "�ӿڲ�����lngOper=" & lngOper & " LngStockID=" & LngStockID & " strNO=" & strNO
    strOutput = strOutput & vbCrLf & "����裺" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "ִ��ʧ�ܣ�HisTransData_DIH"
    Call OutputLog(strOutput)
End Function


