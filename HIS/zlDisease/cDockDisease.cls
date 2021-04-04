VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDockDisease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function InitDockDisease(ByVal lngSys As Long, ByVal cnOracle As ADODB.Connection) As Boolean
    On Error GoTo errH
    glngSys = lngSys: Set gcnOracle = cnOracle
    If zlGetComLib = False Then Exit Function
    InitDockDisease = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function SatisfyCondition(ByVal lngPatiID As Long, ByVal lngPageId As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errHand
    
    strSQL = "Select b.�ļ�id From ������ϼ�¼ A, ��������ǰ�� B,�����ļ��б� C Where a.����id = [1] And a.��ҳid = [2] And a.����id = b.����id AND B.�ļ�ID=C.ID AND C.����=4" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select b.�ļ�id From ������ϼ�¼ A, ��������ǰ�� B,�����ļ��б� C Where a.����id = [1] And a.��ҳid = [2] And a.���id = b.���id AND B.�ļ�ID=C.ID AND C.����=4"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�Ƿ���ϴ�Ⱦ���ϱ�", lngPatiID, lngPageId)
    If rsTemp.RecordCount > 0 Then
        SatisfyCondition = True
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function EditDiseaseDoc(ByVal frmPar As Object, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatFrom As Long, ByVal lngDeptId As Long, ByVal strDiseaseIDs As String, ByVal strDiagnoseIDs As String, Optional ByRef strReason As String) As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHand
    If strDiseaseIDs = "" And strDiagnoseIDs = "" Then Exit Function
 
    Set rsTemp = SatisfyEditDiseaseDoc(lngPatiID, lngPageId, lngPatFrom, lngDeptId, strDiseaseIDs, strDiagnoseIDs)
    EditDiseaseDoc = EditDiseaseReport(frmPar, rsTemp, lngPatiID, lngPageId, lngPatFrom, lngDeptId, strDiseaseIDs, strDiagnoseIDs, strReason)
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function SatisfyEditDiseaseDoc(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatFrom As Long, ByVal lngDeptId As Long, ByVal strDiseaseIDs As String, ByVal strDiagnoseIDs As String) As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    If strDiseaseIDs = "" And strDiagnoseIDs = "" Then
        Exit Function
    End If

    If strDiseaseIDs <> "" Then
        strSQL = " Union Select �ļ�ID From ��������ǰ�� Where ����ID IN (Select Column_Value From Table(f_Num2list([3])))"
    End If
    If strDiagnoseIDs <> "" Then
        strSQL = strSQL & " Union Select �ļ�ID From ��������ǰ�� Where ���ID IN (Select Column_Value From Table(f_Num2list([4])))"
    End If
    strSQL = Mid(strSQL, 8)
    strSQL = "Select /*+ Rule*/ distinct " & vbNewLine & _
            " a.Id, a.����, a.���, a.����, a.����, a.˵��" & vbNewLine & _
            "From �����ļ��б� A ,(" & strSQL & ") B Where A.ID=B.�ļ�ID  And" & vbNewLine & _
            "(a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From ����Ӧ�ÿ��� C Where c.�ļ�id = a.Id And c.����id = [5]))" & vbNewLine & _
            "Order By a.���"
            
    Set SatisfyEditDiseaseDoc = zlDatabase.OpenSQLRecord(strSQL, "CheckDiseaseFile", lngPatiID, lngPageId, strDiseaseIDs, strDiagnoseIDs, lngDeptId)

    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function EditDiseaseReport(ByVal frmPar As Object, ByVal rsTemp As ADODB.Recordset, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatFrom As Long, ByVal lngDeptId As Long, ByVal strDiseaseIDs As String, ByVal strDiagnoseIDs As String, Optional ByRef strReason As String) As Boolean
    Dim lngFileID As Long, strMsg As String
    Dim objDoc As cEPRDocument, objInfection As Object
    Dim strSQL As String
    Dim objDisease As zlRichEPR.cDockDisease
    Dim objDiseaseFileList As Object
    Dim rsDis As ADODB.Recordset
    Dim blnFinded As Boolean
    Dim ObjTabEpr As cTableEPR            '����ʽ�����༭��
    Dim ObjTabEprView As cTableEPR
    
    On Error GoTo errHand
    
    If rsTemp Is Nothing Then
        Exit Function
    ElseIf rsTemp.RecordCount = 0 Then
        Exit Function
    End If

    Do Until rsTemp.EOF
        strMsg = strMsg & vbCrLf & "��" & rsTemp!���� & "��"
        rsTemp.MoveNext
    Loop
    rsTemp.MoveFirst

    If rsTemp.RecordCount = 1 Then
            If zlDatabase.GetPara("��ҳ������༭���ر��濨", glngSys, 1277, "0") = 1 And rsTemp!���� = 4 Then  '��0��ʾ���� =1��ʾ�����༭
            Else
                If MsgBox("���ݲ��˵������Ϣ�����¼���֤�����滹û����д��" & vbCrLf & vbCrLf & Mid(strMsg, 3) & vbCrLf & vbCrLf & "�Ƿ�������д��", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If CheckDisNum(lngPatiID, lngPageId, lngPatFrom, rsDis, glngOpenedID) Then
                        strReason = frmReason.ShowMe(frmPar, Nvl(rsDis!�������˵��))
                        Call SaveReason(strReason, rsDis!id, rsDis!��¼״̬)
                    End If
                    Exit Function
                End If
            End If
Process:    lngFileID = Val(rsTemp!id & "")
        If rsTemp!���� = 4 Then '�±༭��
            Set objInfection = DynamicCreate("zlDisReportCard.clsDisReportCard", "��Ⱦ�����濨", True)
            If Not objInfection Is Nothing Then
                objInfection.Init gcnOracle, glngSys
                objInfection.OpenDoc frmPar, cprEM_����, lngPatiID, lngPageId, lngPatFrom, 0, lngDeptId, lngFileID
            End If
        ElseIf rsTemp!���� = 2 Then '����༭��
            Set ObjTabEprView = New cTableEPR
            Call ObjTabEprView.InitTableEPR(gcnOracle, glngSys, gstrDBOwer)
            
            If Not ObjTabEpr Is Nothing Then
                blnFinded = ObjTabEpr.Showfrm(lngFileID, lngPatiID, lngPageId, lngPatFrom, lngDeptId)
            End If
            If Not blnFinded Then
                Set ObjTabEpr = New cTableEPR
                ObjTabEpr.InitOpenEPR frmPar, cprEM_����, cprET_�������༭, lngFileID, True, 0, lngPatFrom, lngPatiID, lngPageId, 0, lngDeptId, 0, GetPrivFunc(glngSys, IIf(lngPatFrom = cprPF_סԺ, 1251, 1250)), , , , 1
            End If
        Else
            Set objDoc = New cEPRDocument
            Call objDoc.InitEPRDoc(cprEM_����, cprET_�������༭, lngFileID, lngPatFrom, lngPatiID, lngPageId, 0, lngDeptId, 0, False)
            Call objDoc.ShowEPREditor(frmPar, , vbModal)
        End If
    ElseIf rsTemp.RecordCount > 1 Then
        Set objDisease = New zlRichEPR.cDockDisease

        If objDisease.zlGetFrmDiseaseFileList.ShowMe(frmPar, rsTemp, lngFileID) Then
            rsTemp.Filter = "ID=" & lngFileID
            GoTo Process
        Else
            If CheckDisNum(lngPatiID, lngPageId, lngPatFrom, rsDis, glngOpenedID) Then
                strReason = frmReason.ShowMe(frmPar, Nvl(rsDis!�������˵��))
                Call SaveReason(strReason, rsDis!id, rsDis!��¼״̬)
            End If
        End If
        Set objDisease = Nothing
    End If
    EditDiseaseReport = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function RelateFeedback(ByRef frmParent As Object, ByVal lngID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngFrom As Long, ByVal isRelated As Boolean) As Boolean
'���ܣ���Ⱦ�����濨���������Խ��������������ȡ������
'������lngID -����ID
'      lngPatiID -����ID
'      lngPageID -������ҳID���߹Һ�ID
'      lngFrom-������Դ��1-���2-סԺ
'      isRelated  true-������false-ȡ������
    Dim strSQL As String
    Dim rsDisease As ADODB.Recordset
    Dim strIDs As String
    Dim arrayID() As String
    Dim i As Long
    Dim frmFBList As frmFeedBackList

On Error GoTo errHand
  
    If isRelated Then   '����
        If lngFrom = 1 Then
            strSQL = "select rowNum as NO,a.ID,c.���� as ����, a.�Ǽ�ʱ�� from  �������Լ�¼ A ,���˹Һż�¼ B ,���ű� C where A.�ļ�ID is NULL  and A.�Һŵ� = B.NO and A.����ID = B.����ID and A.��¼״̬ <> 3 and A.�Ǽǿ���ID = C.ID  and A.����ID = [1] and B.ID = [2]"
        ElseIf lngFrom = 2 Then
            strSQL = "select rowNum as NO,a.ID ,c.���� as ����,a.�Ǽ�ʱ�� from  �������Լ�¼ A ,���ű� C  where A.�ļ�ID is NULL  and A.��¼״̬ <> 3  and A.�Ǽǿ���ID = C.ID and A.����ID = [1] and A.��ҳID = [2] "
        End If
        Set rsDisease = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�ñ����Ӧ�����Խ��������", lngPatiID, lngPageId)
        If rsDisease.RecordCount = 1 Then
            strSQL = "Zl_�������Լ���¼_Update(2," & rsDisease!id & "," & lngID & ",NULL,NULL,NULL,NULL)"
            Call zlDatabase.ExecuteProcedure(strSQL, "����������������Խ��������")
        ElseIf rsDisease.RecordCount > 1 Then
            Set frmFBList = New frmFeedBackList
            If frmFBList.ShowMe(frmParent, rsDisease, strIDs) Then
                If strIDs <> "" Then
                    arrayID = Split(strIDs, ",")
                    For i = LBound(arrayID) To UBound(arrayID)
                        If Val(arrayID(i)) <> 0 Then
                            strSQL = "Zl_�������Լ���¼_Update(2," & arrayID(i) & "," & lngID & ",NULL,NULL,NULL,NULL)"
                            Call zlDatabase.ExecuteProcedure(strSQL, "����������������Խ��������")
                        End If
                    Next
                End If
            End If
            If Not frmFBList Is Nothing Then
                Unload frmFBList
                Set frmFBList = Nothing
            End If
        End If
    Else 'ȡ������
        strSQL = "Zl_�������Լ���¼_Update(3, NULL " & "," & lngID & ",NULL,NULL,NULL,NULL)"
        Call zlDatabase.ExecuteProcedure(strSQL, "ȡ��������������Խ���������Ĺ���")
    End If
    
    RelateFeedback = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
