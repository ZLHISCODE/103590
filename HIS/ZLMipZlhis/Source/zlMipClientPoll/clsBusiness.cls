VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusiness"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'��������

Private mstrSQL As String
Private mclsDataOracle As zlDataOracle.clsDataOracle
Private mstrTitle As String
Private mstrDbUser As String
Private mrsSQL As ADODB.Recordset
Private mstrError As String
Private mstrOwner As String
Private mblnDBAUser As Boolean

'######################################################################################################################
'�ӿڷ���

'Public Function InitBusiness(ByVal strOwner As String, ByVal strUser As String, ByVal strPassword As String, ByVal strServer As String) As Boolean
Public Function InitBusiness(ByVal objDataOracle As zlDataOracle.clsDataOracle, ByVal strTitle As String, ByVal strOwner As String) As Boolean
    '******************************************************************************************************************
    '���ܣ���ʼ��
    '������
    '���أ�
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    Set mclsDataOracle = objDataOracle
        
    mstrTitle = "��Ϣ������"
    mblnDBAUser = IsDBAUser
    If mblnDBAUser = False Then
        If IsOwner = False Then
            MsgBox "��¼��ֻ��Ϊ���ݿ�DBA���Ʒ�����ߡ�", vbInformation
            Exit Function
        End If
    End If
    mstrOwner = strOwner
    
    InitBusiness = True
    
    Exit Function
    
errHand:
    
End Function

Public Function GetSystem(ByVal blnDBA As Boolean, ByVal strDBUser As String) As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ���ȡInfotable�ṹ
    '������
    '���أ�
    '******************************************************************************************************************
        
    On Error GoTo errHand
    
    If blnDBA = True Then
        mstrSQL = "Select ���,����,������,�汾�� From zlSystems Order By ���"
    Else
        mstrSQL = "Select ���,����,������,�汾�� From zlSystems Where Upper(������) = [1] Order By ���"
    End If
    Set GetSystem = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strDBUser)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetSystemInfo(ByVal strSystems As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    '���û���Ϣ���蹫����������������ʹ��

    mstrSQL = "Select S.*" & _
            " From zlSystems S,(Select Distinct owner From All_Tables Where Table_Name='���ű�') D" & _
            " Where Upper(S.������)=D.Owner And S.��� In (" & strSystems & ") Order by S.���"
    
    Set GetSystemInfo = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
End Function

Public Function GetUserInfo(ByVal strOwner As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mstrSQL = "Select R.*,D.���� as ���ű���,D.���� as ��������,P.���,P.����,P.����" & _
            " From " & strOwner & ".�ϻ���Ա�� U," & strOwner & ".��Ա�� P," & strOwner & ".���ű� D," & strOwner & ".������Ա R" & _
            " Where U.��ԱID = P.ID And R.����ID = D.ID And P.ID=R.��ԱID and U.�û���=USER And (P.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or P.����ʱ�� Is Null) and R.ȱʡ=1"
    
    Set GetUserInfo = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    
End Function

Public Function IsOwner() As Boolean
    '******************************************************************************************************************
    '���ܣ���ȡInfotable�ṹ
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
        
    mstrSQL = "SELECT ��� FROM ZLTOOLS.ZLSYSTEMS WHERE upper(������)=USER"
    Set rsTmp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    
    IsOwner = (rsTmp.BOF = False)
End Function

Public Function GetAccount(ByVal StrComponent As String) As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    mstrSQL = "SELECT Substr(Lpad(���, 5, '0'), 4) ���, ��� ϵͳ, ���� " & _
             " FROM Zlsystems " & _
             " WHERE ��� IN " & _
             "       (SELECT DISTINCT p.ϵͳ " & _
             "        FROM Zlprograms p, " & _
             "             (SELECT ϵͳ, ��� " & _
             "               FROM (SELECT DISTINCT p.ϵͳ, p.���, r.���� AS ��Ȩ " & _
             "                      FROM Zlprogfuncs p, Zlregfunc r " & _
             "                      WHERE Trunc(p.ϵͳ / 100) = r.ϵͳ(+) AND p.��� = r.���(+) AND p.���� = r.����(+) AND " & _
             "                            (EXISTS (SELECT 1 FROM Session_Roles WHERE Role = 'DBA') OR " & _
             "                             p.ϵͳ IN (SELECT ��� FROM Zlsystems WHERE Upper(������) = USER) OR " & _
             "                             p.ϵͳ IN (SELECT ϵͳ FROM Zlrolegrant g, Session_Roles s WHERE g.��ɫ = s.Role)) " & _
             "                      MINUS " & _
             "                      SELECT DISTINCT s.ϵͳ, s.���, r.���� AS ��Ȩ " & _
             "                      FROM Zlprogprivs s, Zlregfunc r " & _
             "                      WHERE Trunc(s.ϵͳ / 100) = r.ϵͳ(+) AND s.��� = r.���(+) AND s.���� = r.����(+) AND " & _
             "                            (EXISTS (SELECT 1 FROM Session_Roles WHERE Role = 'DBA') OR " & _
             "                             s.ϵͳ IN (SELECT ��� FROM Zlsystems WHERE Upper(������) = USER) OR " & _
             "                             s.ϵͳ IN (SELECT ϵͳ FROM Zlrolegrant g, Session_Roles s WHERE g.��ɫ = s.Role)) AND " & _
             "                            s.������ <> USER AND s.���� IN (SELECT Object_Name " & _
             "                                                            FROM User_Objects " & _
             "                                                            WHERE Object_Type IN ('SEQUENCE', 'TABLE', 'VIEW', 'PROCEDURE', 'FUNCTION', 'PACAKEG'))) " & _
             "               WHERE ��Ȩ IS NULL AND ϵͳ IS NULL OR ��Ȩ IS NOT NULL) f"
    mstrSQL = mstrSQL & "       WHERE p.ϵͳ = f.ϵͳ AND p.��� = f.��� AND " & _
             "              Upper(p.����) IN (" & StrComponent & ")) " & _
             " ORDER BY ����, ���"
        
    Set GetAccount = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
End Function

Private Function IsDBAUser() As Boolean
    '******************************************************************************************************************
    '���ܣ���ȡInfotable�ṹ
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    
    mstrSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
    Set rsTmp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    
    IsDBAUser = (rsTmp.BOF = False)
            
End Function

'Public Function GetUserInfo(ByVal strUser As String) As ADODB.Recordset
'    '******************************************************************************************************************
'    '���ܣ���ȡ��ǰ�û��Ļ�����Ϣ
'    '���أ�����Ado��¼��
'    '******************************************************************************************************************
'
'
'    On Error GoTo errHand
'
'    mstrSQL = "Select User,A.Id, A.���, A.����, A.����, B.�û���, C.����id, D.���� As ������, D.���� As ������" & vbNewLine & _
'        "From ��Ա�� A, �ϻ���Ա�� B, ������Ա C, ���ű� D" & vbNewLine & _
'        "Where A.Id = B.��Աid And A.Id = C.��Աid And C.����id = D.Id And B.�û��� = [1] And C.ȱʡ = 1"
'
'    Set GetUserInfo = zlDataBase.OpenSQLRecord(mstrSQL, mstrTitle, strUser)
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    mstrError = Err.Description
'End Function

Public Function BuildMessageRecord(ByRef strMessageKey As String, _
                                    ByVal strItemKey As String, _
                                    ByVal bytTriggerfrequency As Byte, _
                                    Optional strStartDate As String, _
                                    Optional strEndDate As String, _
                                    Optional ByVal rsCondResult As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ���ȡm_Msg_Record����
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intCol As Integer
    Dim strFieldTitle As String
    Dim strFieldValue As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHand
    
    '����Ƿ����������Ϣ
    Select Case bytTriggerfrequency
    Case 1
        mstrSQL = "Select 1 From zltools.zlmip_send_log Where trigger_date Between Trunc(Sysdate) And Sysdate+1 And item_id=[1]"
        Set rsTmp = zlDataBase.OpenSQLRecord(mstrSQL, mstrTitle, strItemKey)
        If rsTmp.BOF = False Then
            BuildMessageRecord = True
            Exit Function
        End If
    Case 2
        mstrSQL = "Select 1 From zltools.zlmip_send_log Where trigger_date Between [1] And [2] And item_id=[3]"
        Set rsTmp = zlDataBase.OpenSQLRecord(mstrSQL, mstrTitle, CDate(strStartDate), CDate(strEndDate), strItemKey)
        If rsTmp.BOF = False Then
            BuildMessageRecord = True
            Exit Function
        End If
    End Select
    
    Call SQLRecord(mrsSQL)
    
    strMessageKey = zlCommFun.GetGUID
    mstrSQL = "zltools.zlMip_PkgRollService.InsertMessage('" & strMessageKey & "','" & strItemKey & "')"
    Call SQLRecordAdd(mrsSQL, mstrSQL)
    
    If Not (rsCondResult Is Nothing) Then
        If Not rsCondResult.EOF Then
            For intCol = 0 To rsCondResult.Fields.Count - 1
                strFieldTitle = rsCondResult.Fields(intCol).Name
                strFieldValue = zlCommFun.Nvl(rsCondResult.Fields(intCol).Value)
                
                mstrSQL = "zltools.zlMip_PkgRollService.InsertMessageParameter('" & strMessageKey & "'," & intCol & ",'" & strFieldTitle & "','" & strFieldValue & "')"
                Call SQLRecordAdd(mrsSQL, mstrSQL)
            Next
        End If
    End If
    
    If SQLRecordExecute(mclsDataOracle, mrsSQL, False) = False Then
        mstrError = Err.Description
    Else
        BuildMessageRecord = True
    End If
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Function

Public Function BuildMessageContent(ByVal strMessageKey As String, ByVal strMessagePara As String) As Boolean
    '******************************************************************************************************************
    '���ܣ���ȡm_Msg_Record����
    '������
    '���أ�
    '******************************************************************************************************************
        
    On Error GoTo errHand
    
    Call SQLRecord(mrsSQL)

    mstrSQL = "zltools.zlMip_PkgRollService.BuildMessageContent('" & strMessageKey & "','" & strMessagePara & "')"
    
    Call SQLRecordAdd(mrsSQL, mstrSQL)
    
    If SQLRecordExecute(mclsDataOracle, mrsSQL, False) = False Then
        mstrError = Err.Description
    Else
        BuildMessageContent = True
    End If
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Function

Public Function RepeatMessageContent(ByVal strMessageKey As String, ByVal strMessagePara As String, ByVal lngRepeatOrder As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ���ȡm_Msg_Record����
    '������
    '���أ�
    '******************************************************************************************************************
        
    On Error GoTo errHand
    
    Call SQLRecord(mrsSQL)

    mstrSQL = "zltools.zlMip_PkgRollService.RepeatMessageContent('" & strMessageKey & "','" & strMessagePara & "'," & lngRepeatOrder & ")"
    
    Call SQLRecordAdd(mrsSQL, mstrSQL)
    
    If SQLRecordExecute(mclsDataOracle, mrsSQL, False) = False Then
        mstrError = Err.Description
    Else
        RepeatMessageContent = True
    End If
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Function

Public Function GetSendMessgeContent(ByVal strKey As String, Optional ByVal lngRepeatOrder As Long = 0) As String
    '******************************************************************************************************************
    '���ܣ���ȡMsgEvent����
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strStyle As String
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim intLoop As Integer
    
    On Error GoTo errHand
    
    mstrSQL = "Select length(a.send_content.getclobval()) As ContentLength From zltools.zlmip_send_log a Where a.id=[1]"
    Set rsTemp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strKey)
    If rsTemp.BOF = False Then
        If rsTemp("ContentLength").Value > 2000 Then
            For intLoop = 1 To rsTemp("ContentLength").Value / 2000 + 1
                
                mstrSQL = "select To_Char(substr(a.send_content.getclobval()," & CDbl(intLoop) * 2000 - 1999 & ",2000)) as send_content from zltools.zlmip_send_log a where a.id = [1] "
                Set rsTemp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strKey)
                If rsTemp.BOF = False Then
                    GetSendMessgeContent = GetSendMessgeContent & zlCommFun.Nvl(rsTemp("send_content").Value)
                End If
            Next
        Else
            mstrSQL = "Select a.send_content.getstringval() As send_content From zltools.zlmip_send_log a Where a.id=[1]"
            Set rsTemp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strKey)
            If rsTemp.BOF = False Then
                GetSendMessgeContent = zlCommFun.Nvl(rsTemp("send_content").Value)
            End If
        End If
    End If
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
'    MsgBox Err.Description
'    Resume
End Function

Public Function GetPollCommonItem() As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSQL As String
        
    On Error GoTo errHand
        
    strSQL = _
        "Select y.Id, y.Trigger_Condition, y.Item_Request, y.Item_Code, y.Trigger_Frequency, x.Start_Date1 As Start_Date," & vbNewLine & _
        "       x.Stop_Date1 As Stop_Date" & vbNewLine & _
        "From (Select *" & vbNewLine & _
        "       From (Select a.Id, a.Start_Date As Start_Date1, Nvl(a.Stop_Date, To_Date('3000-01-01', 'yyyy-mm-dd')) As Stop_Date1" & vbNewLine & _
        "              From Zlmip_Item A" & vbNewLine & _
        "              Where a.Trigger_Type = 2 And a.Check_Frequency = 0)" & vbNewLine & _
        "       Where Sysdate Between Start_Date1 And Nvl(Stop_Date1, To_Date('3000-01-01', 'yyyy-mm-dd'))"
        
    'ÿ����
    strSQL = strSQL & _
            "Union All" & vbNewLine & _
            "Select *" & vbNewLine & _
            "From (Select a.Id," & vbNewLine & _
            "              To_Date(To_Char(Decode(a.Check_Freq_Internal, 1, Sysdate, Nvl(a.Last_Date, Sysdate)), 'yyyy-mm-dd') || ' ' ||" & vbNewLine & _
            "                       b.Freq_Start || ':00', 'yyyy-mm-dd hh24:mi:ss') As Start_Date1," & vbNewLine & _
            "              To_Date(To_Char(Decode(a.Check_Freq_Internal, 1, Sysdate, Nvl(a.Last_Date, Sysdate)), 'yyyy-mm-dd') || ' ' ||" & vbNewLine & _
            "                       b.Freq_Stop || ':59', 'yyyy-mm-dd hh24:mi:ss') As Stop_Date1" & vbNewLine & _
            "       From Zlmip_Item A, Zlmip_Item_Frequency B" & vbNewLine & _
            "       Where a.Trigger_Type = 2 And a.Check_Frequency = 1 And a.Id = b.Item_Id And Sysdate Between a.Start_Date And Nvl(a.Stop_Date, To_Date('3000-01-01', 'yyyy-mm-dd')))" & vbNewLine & _
            "Where Sysdate Between Start_Date1 And Nvl(Stop_Date1, To_Date('3000-01-01', 'yyyy-mm-dd'))"
    
    'ÿ�ܼ��
    strSQL = strSQL & _
            "Union All" & vbNewLine & _
            "Select *" & vbNewLine & _
            "From (Select a.Id," & vbNewLine & _
            "              To_Date(To_Char(Next_Day(Decode(a.Check_Freq_Internal, 1, Sysdate, Nvl(a.Last_Date, Sysdate)) - 7," & vbNewLine & _
            "                                        '������') + 7 * (a.Check_Freq_Internal - 1) + b.Freq_Day, 'yyyy-mm-dd') || ' ' ||" & vbNewLine & _
            "                       b.Freq_Start || ':00', 'yyyy-mm-dd hh24:mi:ss') As Start_Date1," & vbNewLine & _
            "              To_Date(To_Char(Next_Day(Decode(a.Check_Freq_Internal, 1, Sysdate, Nvl(a.Last_Date, Sysdate)) - 7," & vbNewLine & _
            "                                        '������') + 7 * (a.Check_Freq_Internal - 1) + b.Freq_Day, 'yyyy-mm-dd') || ' ' ||" & vbNewLine & _
            "                       b.Freq_Stop || ':59', 'yyyy-mm-dd hh24:mi:ss') As Stop_Date1" & vbNewLine & _
            "       From Zlmip_Item A, Zlmip_Item_Frequency B" & vbNewLine & _
            "       Where a.Trigger_Type = 2 And a.Check_Frequency = 2 And a.Id = b.Item_Id And Sysdate Between a.Start_Date And Nvl(a.Stop_Date, To_Date('3000-01-01', 'yyyy-mm-dd')))" & vbNewLine & _
            "Where Sysdate Between Start_Date1 And Nvl(Stop_Date1, To_Date('3000-01-01', 'yyyy-mm-dd'))"
    
    'ÿ�¼��
    strSQL = strSQL & _
            "Union All" & vbNewLine & _
            "Select *" & vbNewLine & _
            "From (Select a.Id," & vbNewLine & _
            "              To_Date(To_Char(Add_Months(Decode(a.Check_Freq_Internal, 1, Sysdate, Nvl(a.Last_Date, Sysdate))," & vbNewLine & _
            "                                          a.Check_Freq_Internal - 1), 'yyyy-mm') || '-' || b.Freq_Day || ' ' ||" & vbNewLine & _
            "                       b.Freq_Start || ':00', 'yyyy-mm-dd hh24:mi:ss') As Start_Date1," & vbNewLine & _
            "              To_Date(To_Char(Add_Months(Decode(a.Check_Freq_Internal, 1, Sysdate, Nvl(a.Last_Date, Sysdate))," & vbNewLine & _
            "                                          a.Check_Freq_Internal - 1), 'yyyy-mm') || '-' || b.Freq_Day || ' ' ||" & vbNewLine & _
            "                       b.Freq_Stop || ':59', 'yyyy-mm-dd hh24:mi:ss') As Stop_Date1" & vbNewLine & _
            "       From Zlmip_Item A, Zlmip_Item_Frequency B" & vbNewLine & _
            "       Where a.Trigger_Type = 2 And a.Check_Frequency = 3 And a.Id = b.Item_Id And Sysdate Between a.Start_Date And Nvl(a.Stop_Date, To_Date('3000-01-01', 'yyyy-mm-dd')))" & vbNewLine & _
            "Where Sysdate Between Start_Date1 And Nvl(Stop_Date1, To_Date('3000-01-01', 'yyyy-mm-dd'))"
        
    'ÿ����
    strSQL = strSQL & _
        "       Union All" & vbNewLine & _
        "       Select *" & vbNewLine & _
        "       From (Select a.Id," & vbNewLine & _
        "                     To_Date(To_Char(To_Number(To_Char(Decode(a.Check_Freq_Internal, 1, Sysdate, Nvl(a.Last_Date, Sysdate))," & vbNewLine & _
        "                                                        'yyyy')) + a.Check_Freq_Internal - 1) || '-' ||" & vbNewLine & _
        "                              Nvl(b.Freq_Month, '1') || '-' || b.Freq_Day || ' ' || b.Freq_Start || ':00'," & vbNewLine & _
        "                              'yyyy-mm-dd hh24:mi:ss') As Start_Date1," & vbNewLine & _
        "                     To_Date(To_Char(To_Number(To_Char(Decode(a.Check_Freq_Internal, 1, Sysdate, Nvl(a.Last_Date, Sysdate))," & vbNewLine & _
        "                                                        'yyyy')) + a.Check_Freq_Internal - 1) || '-' ||" & vbNewLine & _
        "                              Nvl(b.Freq_Month, '1') || '-' || b.Freq_Day || ' ' || b.Freq_Stop || ':59'," & vbNewLine & _
        "                              'yyyy-mm-dd hh24:mi:ss') As Stop_Date1" & vbNewLine & _
        "              From Zlmip_Item A, Zlmip_Item_Frequency B" & vbNewLine & _
        "              Where a.Trigger_Type = 2 And a.Check_Frequency = 4 And a.Id = b.Item_Id And Sysdate Between a.Start_Date And Nvl(a.Stop_Date, To_Date('3000-01-01', 'yyyy-mm-dd')))" & vbNewLine & _
        "       Where Sysdate Between Start_Date1 And Nvl(Stop_Date1, To_Date('3000-01-01', 'yyyy-mm-dd'))) X, Zlmip_Item Y" & vbNewLine & _
        "Where x.Id = y.Id "
            
    If mblnDBAUser = False Then strSQL = strSQL & " And Y.owner=[1]"
    
    Set GetPollCommonItem = zlDataBase.OpenSQLRecord(strSQL, mstrTitle, UCase(mstrOwner))
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetPollAgainMessage() As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSQL As String
        
    On Error GoTo errHand
    
    strSQL = _
        "Select y.Id,z.item_request,z.item_code,Nvl(Max(x.send_order),0) As repeatorder" & vbNewLine & _
        "From (Select a.Id" & vbNewLine & _
        "       From zlmip_Send_Log A" & vbNewLine & _
        "       Where a.Send_Again = 1" & vbNewLine & _
        "       Union All" & vbNewLine & _
        "       Select a.Id" & vbNewLine & _
        "       From zlmip_Send_Log A" & vbNewLine & _
        "       Where a.Send_Again = 2 And a.Send_Again_Number < To_Number(a.Send_Again_Para)" & vbNewLine & _
        "       Union All" & vbNewLine & _
        "       Select a.Id" & vbNewLine & _
        "       From zlmip_Send_Log A" & vbNewLine & _
        "       Where a.Send_Again = 3 And a.Send_Date + To_Number(a.Send_Again_Para)<Sysdate) X," & vbNewLine & _
        "     zlmip_Send_Log Y,zlmip_item Z,zlmip_sendlog_again X" & vbNewLine & _
        "Where x.Id = y.Id And Y.item_identity=Z.item_code And x.send_log_id(+)=y.id Group By y.Id,z.item_request,z.item_code"
            
    Set GetPollAgainMessage = zlDataBase.OpenSQLRecord(strSQL, mstrTitle)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetOptionsValue(ByVal lngParaNo As Long) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ����ֵ From zlOptions Where ������=[1]"
    Set rsTmp = zlDataBase.OpenSQLRecord(strSQL, mstrTitle, lngParaNo)
    If rsTmp.BOF = False Then
        GetOptionsValue = zlCommFun.Nvl(rsTmp("����ֵ").Value)
    End If
        
End Function
    
Public Function UpdateOptionsValue(ByVal lngParaNo As Long, ByVal strParaValue As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim strSQL As String
    
    strSQL = "Update zlOptions Set ����ֵ='" & strParaValue & "' where ������=" & lngParaNo
    Call gclsMsgOracle.ExecuteSQL(strSQL, mstrTitle)
            
    UpdateOptionsValue = True
    
End Function

Public Function SendLogEdit(ByVal strSendLogKey As String, ByVal bytSendState As Byte, ByVal strSendmspUser As String, ByVal strSendUser As String, ByVal strSendUsername As String, ByVal strSendStation As String, Optional ByVal strSendNote As String) As Boolean
    '******************************************************************************************************************
    '���ܣ���ȡm_Msg_Record����
    '������
    '���أ�
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    Call SQLRecord(mrsSQL)
    
    mstrSQL = "zltools.zlMip_PkgRollService.UpdateSendLogState('" & strSendLogKey & "'," & bytSendState & ",'" & strSendmspUser & "','" & strSendUser & "','" & strSendUsername & "','" & strSendStation & "','" & strSendNote & "')"
    
    Call SQLRecordAdd(mrsSQL, mstrSQL)
    
    If SQLRecordExecute(mclsDataOracle, mrsSQL, False) = False Then
        mstrError = Err.Description
    Else
        SendLogEdit = True
    End If
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Function

Public Function SQLRecordExecute(ByVal objDataOracle As clsDataOracle, ByVal rsSQL As ADODB.Recordset, Optional ByVal blnHaveTrans As Boolean = True, Optional ByRef strError As String) As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim blnTran As Boolean
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim strTemp As String
    Dim aryTemp As Variant
    Dim strSQL As String
    
    On Error GoTo errHand
        
    If rsSQL.RecordCount > 0 Then

        blnTran = True

        If blnHaveTrans Then objDataOracle.BeginTrans

        rsSQL.MoveFirst

        For intLoop = 1 To rsSQL.RecordCount

            strSQL = CStr(rsSQL("SQL").Value)
            Call zlDataBase.ExecuteProcedure(strSQL, "")

            rsSQL.MoveNext
        Next

        If blnHaveTrans Then objDataOracle.CommitTrans
        blnTran = False
    End If
    
    SQLRecordExecute = True
    
    Exit Function
errHand:
    
    If blnTran And blnHaveTrans Then objDataOracle.RollbackTrans
    strError = Err.Description
'    MsgBox Err.Description
    
'    If ErrCenter = 1 Then
'        Resume
'    End If
        
End Function

Private Sub Class_Terminate()
    
    Set mclsDataOracle = Nothing
End Sub