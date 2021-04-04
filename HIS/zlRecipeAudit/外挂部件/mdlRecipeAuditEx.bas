Attribute VB_Name = "mdlRecipeAuditEx"
Option Explicit

Public gcnOracle As ADODB.Connection
Public gbytClass As Byte
Public gblnInit As Boolean
Public grsCheckItems As ADODB.Recordset
Public gstrIDs As String

'此编码要与“处方审查项目”表的编码值相同，否则无法调用对应的功能函数
Public Const GSTR_CODE_中药注射剂  As String = "D01"
'Public Const GSTR_CODE_新审查项目 As String = "...."

Public Function F_中药注射剂(ByRef strMedicalID As String, ByRef strErr As String) As Boolean
'功能：审查传入的一批医嘱中，是否存在两种（含两种）以上的中药注射剂
'参数：
'  strMedicalID（实参）：合格/不合格的医嘱ID
'  strErr（实参）：异常信息
'返回：True合格；False不合格

    Dim i As Integer, j As Integer
    Dim arrTmp As Variant
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, strIDs As String
    Dim intCount As Integer
    
    strIDs = gstrIDs
    
    arrTmp = Split(strIDs, ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        If Trim(arrTmp(i)) <> "" Then j = j + 1
    Next
    If j = 1 Then
        F_中药注射剂 = True
        Exit Function
    End If
    
    On Error GoTo errHandle
    
    '检查中药注射剂医嘱
    If Right(strIDs, 1) = "," Then strIDs = Left(strIDs, Len(strIDs) - 1)

    '当前提交的药嘱是否存在两种和两种以上
    strSQL = "Select a.ID " & _
             "From 病人医嘱记录 A, 诊疗项目目录 B, 药品特性 C, Table(f_Num2list([1], ',')) D " & _
             "Where a.相关Id = d.Column_Value And a.诊疗项目id = b.Id And b.Id = c.药名id And b.类别 = '6' " & _
             "  And c.药品剂型 Like '%注射剂%' And ROWNUM < 3 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "验证中药注射剂1", strIDs)
    If rsTemp.RecordCount >= 2 Then
        '中药注射剂医嘱大于一项
        rsTemp.MoveLast
        strMedicalID = CStr(rsTemp!ID)
        F_中药注射剂 = False
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    
    '再检查24小时内的药嘱是否存在两种和两种以上
'    strSQL = "Select a.id " & vbNewLine & _
'             "From 病人医嘱记录 A, 诊疗项目目录 B, 药品特性 C," & vbNewLine & _
'             "     (Select a.病人id " & vbNewLine & _
'             "      From 病人医嘱记录 A, Table(f_Num2list([1], ',') ) B " & vbNewLine & _
'             "      Where a.相关id = b.Column_Value And Rownum < 2) D " & vbNewLine & _
'             "Where a.病人id = d.病人id And a.诊疗项目id = b.Id And b.Id = c.药名id And b.类别 = '6' And c.药品剂型 Like '%注射剂%' " & vbNewLine & _
'             "    And a.开嘱时间 >= Sysdate - 1 And Rownum < 3 "
    strSQL = "Select a.Id " & vbNewLine & _
             "From 病人医嘱记录 A, 诊疗项目目录 B, 药品特性 C, 病人医嘱记录 D, Table(f_Num2list([1], ',')) E " & vbNewLine & _
             "Where a.病人id = d.病人id And d.相关id = e.Column_Value And a.诊疗项目id = b.Id And b.Id = c.药名id And b.类别 = '6' " & vbNewLine & _
             "    And c.药品剂型 Like '%注射剂%' And a.开嘱时间 >= Sysdate - 1 And Rownum < 3 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "验证中药注射剂2", strIDs)
    If rsTemp.RecordCount >= 2 Then
        rsTemp.MoveLast
        strMedicalID = CStr(rsTemp!ID)
        F_中药注射剂 = False
    Else
        strMedicalID = ""
        F_中药注射剂 = True
    End If
    rsTemp.Close
    
    Exit Function
    
errHandle:
    strMedicalID = ""
    If zl9ComLib.ErrCenter() = 1 Then
        Resume
    Else
        strErr = Err.Description
    End If
End Function

'Public Function F_新方法(...) As Boolean
''功能：
''参数：
''返回：
'End Function
