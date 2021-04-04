Attribute VB_Name = "mdlPrint"
Option Explicit

Public Function ReportPrint(ByVal lngKey As Long, ByVal frmPrint As Form, ByVal objPrint As Object, ByVal blnPrint As Boolean) As String
        '功能:  单个报告打印
        '参数:
        '       lngKey          标本ID
        '
        '       blnPrint        True=打印
        Dim strReportCode As String         '报表格式编号
        Dim strReportParaNo As String       '申请单号
        Dim bytReportParaMode As Byte       '医嘱记录性质
        Dim rsTmp As New ADODB.Recordset
        Dim blnCurrMoved As Boolean         '数据是否已转移
        Dim strSQL As String
        Dim strChart(0 To 9) As String
        Dim intLoop As Integer
        Dim strErr As String
        Dim lngAdviceID As Long, lngPatiID As Long
        On Error GoTo errH
    
     
         '生成图形供自定义报表调用
100     strSQL = "Select ID,医嘱ID,病人ID from 检验标本记录 where id = [1] "
102     Set rsTmp = ComOpenSQL(strSQL, "ReportPrint", lngKey)
    
104     Do Until rsTmp.EOF
106         lngAdviceID = Val("" & rsTmp!医嘱ID)
108         lngPatiID = Val("" & rsTmp!病人ID)
        
110         If ReadSampleImage(lngKey, strChart, strErr) = False Then
                '失败，不提示，就是没图形
112             If strErr <> "" Then ShowLog LOG_PRINTSVR, LOG_WARNING, "打印报告", 100, "读取图形失败！" & strErr
            End If
        
114         If GetReportCode(lngAdviceID, 0, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
                If strReportCode = "" Then
                    ReportPrint = "检验报告格式未设定！"
                    Exit Function
                Else
116             Call objPrint.ReportOpen(gcnOracle, 100, strReportCode, frmPrint, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, "医嘱ID=" & lngAdviceID, _
                                "病人ID=" & lngPatiID, "标本ID=" & lngKey, "多个医嘱=" & lngAdviceID, "多个标本=" & lngKey, _
                                "图形1=" & strChart(0), "图形2=" & strChart(1), "图形3=" & strChart(2), "图形4=" & strChart(3), _
                                "图形5=" & strChart(4), "图形6=" & strChart(5), "图形7=" & strChart(6), "图形8=" & strChart(7), _
                                "图形9=" & strChart(8), IIf(blnPrint, 2, 1))
                End If
            End If
    
    
            On Error GoTo errH

118         strSQL = "ZL_检验标本记录_标本质控(" & rsTmp("ID") & ",'',1)"
120         ComExecuteProc strSQL, "打印"
122         rsTmp.MoveNext
        Loop
        ReportPrint = "OK"
        On Error Resume Next
        '删除图形文件
126     For intLoop = 1 To 9
128         Kill strChart(intLoop)
        Next
130
        Exit Function
errH:
132     ReportPrint = "ReportPrint " & CStr(Erl()) & "行，" & Err.Description
134     ShowLog LOG_PRINTSVR, LOG_ERR, "打印报告", Err.Number, ReportPrint
End Function

Private Function GetReportCode(ByVal lngAdviceID As Long, ByVal lng发送号 As Long, _
                               ByRef strCode As String, ByRef strNo As String, _
                               ByRef bytMode As Byte, Optional ByVal DataMoved As Boolean = False) As Boolean
    
        '功能;  获取报表编号
    
        Dim rs As New ADODB.Recordset
        Dim strSQL As String
    
        On Error GoTo errH
    
100     If lngAdviceID = 0 And lng发送号 = 0 Then Exit Function
    
    '    strSQL = "SELECT DISTINCT 'ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2' AS 报表编号," & _
                           "A.NO," & _
                           "A.记录性质 " & _
                    "FROM 病人医嘱发送 A,病历文件列表 C,病人医嘱记录 D,病历单据应用 E " & _
                    "Where E.病历文件id = C.ID " & _
                            "AND D.诊疗项目ID=E.诊疗项目ID " & _
                          "AND A.医嘱ID=D.ID AND E.应用场合=Decode(D.病人来源,2,2,4,4,1) " & _
                          " AND D.相关id= [1] "
                      
102     strSQL = "Select Distinct 'ZLCISBILL' || Trim(To_Char(C.编号, '00000')) || '-2' As 报表编号, A.NO, Nvl(A.记录性质,1) as 记录性质, F.ID, F.编码" & vbNewLine & _
                "From 病人医嘱发送 A, 病历文件列表 C, 病人医嘱记录 D, 病历单据应用 E, 诊疗项目目录 F" & vbNewLine & _
                "Where E.病历文件id = C.ID And D.诊疗项目id = E.诊疗项目id And D.诊疗项目id = F.ID And A.医嘱id = D.ID And" & vbNewLine & _
                "      E.应用场合 = Decode(D.病人来源, 2, 2, 4, 4, 1) And D.相关id = [1] " & vbNewLine & _
                "Order By F.编码 "
                          
104     If DataMoved Then
106         strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
108         strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        End If

110     Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlLISWork", lngAdviceID, lng发送号)
                      
    
112     If rs.BOF = False Then
114         strCode = Trim("" & rs("报表编号"))
116         strNo = Trim("" & rs("NO"))
118         bytMode = Val("" & rs("记录性质"))
        End If
    
120     GetReportCode = True
        Exit Function
errH:
122     ShowLog LOG_PRINTSVR, LOG_ERR, "取报表编号", Err.Number, CStr(Erl()) & "行," & Err.Description

End Function
