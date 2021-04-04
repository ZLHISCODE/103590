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
'变量定义

Private mstrSQL As String
Private mclsDataOracle As zlDataOracle.clsDataOracle
Private mstrTitle As String
Private mstrDbUser As String
Private mrsSQL As ADODB.Recordset
Private mstrError As String
Private mstrOwner As String
Private mblnDBAUser As Boolean

'######################################################################################################################
'接口方法

'Public Function InitBusiness(ByVal strOwner As String, ByVal strUser As String, ByVal strPassword As String, ByVal strServer As String) As Boolean
Public Function InitBusiness(ByVal objDataOracle As zlDataOracle.clsDataOracle, ByVal strTitle As String, ByVal strOwner As String) As Boolean
    '******************************************************************************************************************
    '功能：初始化
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    Set mclsDataOracle = objDataOracle
        
    mstrTitle = "消息发送器"
    mblnDBAUser = IsDBAUser
    If mblnDBAUser = False Then
        If IsOwner = False Then
            MsgBox "登录者只能为数据库DBA或产品所有者。", vbInformation
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
    '功能：读取Infotable结构
    '参数：
    '返回：
    '******************************************************************************************************************
        
    On Error GoTo errHand
    
    If blnDBA = True Then
        mstrSQL = "Select 编号,名称,所有者,版本号 From zlSystems Order By 编号"
    Else
        mstrSQL = "Select 编号,名称,所有者,版本号 From zlSystems Where Upper(所有者) = [1] Order By 编号"
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
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    '读用户信息赋予公共，便于其他程序使用

    mstrSQL = "Select S.*" & _
            " From zlSystems S,(Select Distinct owner From All_Tables Where Table_Name='部门表') D" & _
            " Where Upper(S.所有者)=D.Owner And S.编号 In (" & strSystems & ") Order by S.编号"
    
    Set GetSystemInfo = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
End Function

Public Function GetUserInfo(ByVal strOwner As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mstrSQL = "Select R.*,D.编码 as 部门编码,D.名称 as 部门名称,P.编号,P.姓名,P.简码" & _
            " From " & strOwner & ".上机人员表 U," & strOwner & ".人员表 P," & strOwner & ".部门表 D," & strOwner & ".部门人员 R" & _
            " Where U.人员ID = P.ID And R.部门ID = D.ID And P.ID=R.人员ID and U.用户名=USER And (P.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or P.撤档时间 Is Null) and R.缺省=1"
    
    Set GetUserInfo = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    
End Function

Public Function IsOwner() As Boolean
    '******************************************************************************************************************
    '功能：读取Infotable结构
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
        
    mstrSQL = "SELECT 编号 FROM ZLTOOLS.ZLSYSTEMS WHERE upper(所有者)=USER"
    Set rsTmp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    
    IsOwner = (rsTmp.BOF = False)
End Function

Public Function GetAccount(ByVal StrComponent As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    mstrSQL = "SELECT Substr(Lpad(编号, 5, '0'), 4) 编号, 编号 系统, 名称 " & _
             " FROM Zlsystems " & _
             " WHERE 编号 IN " & _
             "       (SELECT DISTINCT p.系统 " & _
             "        FROM Zlprograms p, " & _
             "             (SELECT 系统, 序号 " & _
             "               FROM (SELECT DISTINCT p.系统, p.序号, r.功能 AS 授权 " & _
             "                      FROM Zlprogfuncs p, Zlregfunc r " & _
             "                      WHERE Trunc(p.系统 / 100) = r.系统(+) AND p.序号 = r.序号(+) AND p.功能 = r.功能(+) AND " & _
             "                            (EXISTS (SELECT 1 FROM Session_Roles WHERE Role = 'DBA') OR " & _
             "                             p.系统 IN (SELECT 编号 FROM Zlsystems WHERE Upper(所有者) = USER) OR " & _
             "                             p.系统 IN (SELECT 系统 FROM Zlrolegrant g, Session_Roles s WHERE g.角色 = s.Role)) " & _
             "                      MINUS " & _
             "                      SELECT DISTINCT s.系统, s.序号, r.功能 AS 授权 " & _
             "                      FROM Zlprogprivs s, Zlregfunc r " & _
             "                      WHERE Trunc(s.系统 / 100) = r.系统(+) AND s.序号 = r.序号(+) AND s.功能 = r.功能(+) AND " & _
             "                            (EXISTS (SELECT 1 FROM Session_Roles WHERE Role = 'DBA') OR " & _
             "                             s.系统 IN (SELECT 编号 FROM Zlsystems WHERE Upper(所有者) = USER) OR " & _
             "                             s.系统 IN (SELECT 系统 FROM Zlrolegrant g, Session_Roles s WHERE g.角色 = s.Role)) AND " & _
             "                            s.所有者 <> USER AND s.对象 IN (SELECT Object_Name " & _
             "                                                            FROM User_Objects " & _
             "                                                            WHERE Object_Type IN ('SEQUENCE', 'TABLE', 'VIEW', 'PROCEDURE', 'FUNCTION', 'PACAKEG'))) " & _
             "               WHERE 授权 IS NULL AND 系统 IS NULL OR 授权 IS NOT NULL) f"
    mstrSQL = mstrSQL & "       WHERE p.系统 = f.系统 AND p.序号 = f.序号 AND " & _
             "              Upper(p.部件) IN (" & StrComponent & ")) " & _
             " ORDER BY 名称, 编号"
        
    Set GetAccount = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
End Function

Private Function IsDBAUser() As Boolean
    '******************************************************************************************************************
    '功能：读取Infotable结构
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    
    mstrSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
    Set rsTmp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    
    IsDBAUser = (rsTmp.BOF = False)
            
End Function

'Public Function GetUserInfo(ByVal strUser As String) As ADODB.Recordset
'    '******************************************************************************************************************
'    '功能：获取当前用户的基本信息
'    '返回：返回Ado记录集
'    '******************************************************************************************************************
'
'
'    On Error GoTo errHand
'
'    mstrSQL = "Select User,A.Id, A.编号, A.简码, A.姓名, B.用户名, C.部门id, D.编码 As 部门码, D.名称 As 部门名" & vbNewLine & _
'        "From 人员表 A, 上机人员表 B, 部门人员 C, 部门表 D" & vbNewLine & _
'        "Where A.Id = B.人员id And A.Id = C.人员id And C.部门id = D.Id And B.用户名 = [1] And C.缺省 = 1"
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
    '功能：读取m_Msg_Record数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intCol As Integer
    Dim strFieldTitle As String
    Dim strFieldValue As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHand
    
    '检查是否可以生成消息
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
    '功能：读取m_Msg_Record数据
    '参数：
    '返回：
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
    '功能：读取m_Msg_Record数据
    '参数：
    '返回：
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
    '功能：读取MsgEvent数据
    '参数：
    '返回：
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
    '功能：
    '参数：
    '返回：
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
        
    '每天检查
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
    
    '每周检查
    strSQL = strSQL & _
            "Union All" & vbNewLine & _
            "Select *" & vbNewLine & _
            "From (Select a.Id," & vbNewLine & _
            "              To_Date(To_Char(Next_Day(Decode(a.Check_Freq_Internal, 1, Sysdate, Nvl(a.Last_Date, Sysdate)) - 7," & vbNewLine & _
            "                                        '星期日') + 7 * (a.Check_Freq_Internal - 1) + b.Freq_Day, 'yyyy-mm-dd') || ' ' ||" & vbNewLine & _
            "                       b.Freq_Start || ':00', 'yyyy-mm-dd hh24:mi:ss') As Start_Date1," & vbNewLine & _
            "              To_Date(To_Char(Next_Day(Decode(a.Check_Freq_Internal, 1, Sysdate, Nvl(a.Last_Date, Sysdate)) - 7," & vbNewLine & _
            "                                        '星期日') + 7 * (a.Check_Freq_Internal - 1) + b.Freq_Day, 'yyyy-mm-dd') || ' ' ||" & vbNewLine & _
            "                       b.Freq_Stop || ':59', 'yyyy-mm-dd hh24:mi:ss') As Stop_Date1" & vbNewLine & _
            "       From Zlmip_Item A, Zlmip_Item_Frequency B" & vbNewLine & _
            "       Where a.Trigger_Type = 2 And a.Check_Frequency = 2 And a.Id = b.Item_Id And Sysdate Between a.Start_Date And Nvl(a.Stop_Date, To_Date('3000-01-01', 'yyyy-mm-dd')))" & vbNewLine & _
            "Where Sysdate Between Start_Date1 And Nvl(Stop_Date1, To_Date('3000-01-01', 'yyyy-mm-dd'))"
    
    '每月检查
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
        
    '每年检查
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
    '功能：
    '参数：
    '返回：
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
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 参数值 From zlOptions Where 参数号=[1]"
    Set rsTmp = zlDataBase.OpenSQLRecord(strSQL, mstrTitle, lngParaNo)
    If rsTmp.BOF = False Then
        GetOptionsValue = zlCommFun.Nvl(rsTmp("参数值").Value)
    End If
        
End Function
    
Public Function UpdateOptionsValue(ByVal lngParaNo As Long, ByVal strParaValue As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim strSQL As String
    
    strSQL = "Update zlOptions Set 参数值='" & strParaValue & "' where 参数号=" & lngParaNo
    Call gclsMsgOracle.ExecuteSQL(strSQL, mstrTitle)
            
    UpdateOptionsValue = True
    
End Function

Public Function SendLogEdit(ByVal strSendLogKey As String, ByVal bytSendState As Byte, ByVal strSendmspUser As String, ByVal strSendUser As String, ByVal strSendUsername As String, ByVal strSendStation As String, Optional ByVal strSendNote As String) As Boolean
    '******************************************************************************************************************
    '功能：读取m_Msg_Record数据
    '参数：
    '返回：
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
    '功能:
    '参数:
    '返回:
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
