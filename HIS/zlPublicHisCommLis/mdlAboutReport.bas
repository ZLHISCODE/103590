Attribute VB_Name = "mdlAboutReport"
Option Explicit

Public Function findThirdReport(ByVal lngSampleID As String, objWeb As WebBrowser)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsTemp As Recordset
    '三方LIS报告
    Dim strTag As String

    strSQL = "select 医嘱ID,申请ID from 检验申请组合 where 标本ID=[1] and 医嘱ID is not null"
    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验申请组合", lngSampleID)
    Do While Not rsTmp.EOF
        strSQL = "select b.id as 报告ID,b.报告名,b.报告名||','||To_Char(A.开嘱时间,'YYYY-MM-DD HH24:MI') as 文档标题,c.医嘱ID,b.类型,b.打印次数 from 病人医嘱记录 a, 医嘱报告内容 b,病人医嘱报告 c where b.id=c.报告id and a.id=c.医嘱id and c.报告id is not null and b.类型 in (0,2) and a.id =[1]" & vbCrLf & _
               " union all " & vbCrLf & _
               " select b.id as 报告ID,b.报告名,b.报告名||','||To_Char(A.开嘱时间,'YYYY-MM-DD HH24:MI') as 文档标题,c.医嘱ID,b.类型,b.打印次数 from 病人医嘱记录 a, 医嘱报告内容 b,病人医嘱报告 c where b.id=c.报告id and a.id=c.医嘱id and c.报告id is not null and b.类型 in (0,2) and a.id =[2]"

        Set rsTemp = OpenSQLRecord(Sel_His_DB, strSQL, "三方报告", Val(rsTmp("医嘱ID") & ""), Val(rsTmp("申请ID") & ""))
        If rsTemp.RecordCount > 0 Then
            strTag = strTag & "<SP>" & rsTemp!报告ID & ";" & rsTemp!医嘱id & ";" & rsTemp!类型 & "<sTab>" & rsTemp!报告名
            Call WebShow(strTag, objWeb)
        End If
    Loop
    If strTag <> "" Then findThirdReport = Mid(strTag, 5)

End Function

Public Sub WebShow(ByVal strKey As String, objWeb As WebBrowser)
'功能：Web控件展示文件
    Dim strURL As String
    If strKey = "" Then
        Call objWeb.Navigate("about:blank")
        objWeb.Visible = False
'        mstrCurFile = ""
    Else
        strURL = GetLisRptFile(strKey)
        If strURL <> "" Then
            objWeb.Navigate strURL
'            mstrCurFile = strURL
        End If
        objWeb.Visible = True
    End If
End Sub

Public Function GetLisRptFile(ByVal strTag As String) As String
'功能：打开LIS报告文件查看，获取临时文件路径
    Dim strFile As String
    Dim objFile As New FileSystemObject
    Dim strTmp As String
    Dim lng报告ID As String
    Dim str报告名 As String
    Dim lng类型 As String
    Dim varTmp As Variant
    Dim strSuffix As String '文件后缀名
    
    Screen.MousePointer = 11
    
    varTmp = Split(strTag, ";")
    lng报告ID = varTmp(0)
    strTmp = Replace(strTag, varTmp(0) & ";" & varTmp(1) & ";", "")
    varTmp = Split(strTmp, "<sTab>")
    lng类型 = varTmp(0)
    If lng类型 = 0 Then
        strSuffix = "pdf"
    ElseIf lng类型 = 1 Then
        strSuffix = "html"
    Else
        strSuffix = "xps"
    End If
    str报告名 = varTmp(1)
    
    strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\tmpReport_" & lng报告ID & "." & strSuffix
    If Not objFile.FileExists(strFile) Then
        strFile = ReadLob(100, 22, lng报告ID, strFile)
        If Not objFile.FileExists(strFile) Then
            MsgBox "文件内容读取失败！", vbInformation, "中联信息":
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    GetLisRptFile = strFile
    Screen.MousePointer = 0
End Function

Public Function CalcVolatility(strCalcA As String, strCalcB As String) As String
    '计算变异率

    On Error Resume Next

    If strCalcA = "" Or strCalcB = "" Then
        CalcVolatility = ""
        Exit Function
    End If
    If Val(strCalcA) = 0 Or Val(strCalcB) = 0 Then
        CalcVolatility = ""
    End If

    '计算
    CalcVolatility = (Val(strCalcB) - Val(strCalcA)) / Val(strCalcA) * 100
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/8/27
'功    能:动态创建控件
'入    参:
'           objParent           父级对象，需要在哪个窗体中创建对象
'           strControlClass     需要创建的控件
'           strControlClass     控件名称
'           [objPart            容器对象，要将改控件创建的哪个容器中,默认为父级窗体对象中]
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function NewControl(objParent As Object, ByVal strControlClass As String, ByVal strName As String, Optional objPart As Object) As Object
          Dim objCrl As Object
          
          '加入协议，只能加一次，第二次会出错
          

1         On Error Resume Next
2         Call Licenses.Add(strControlClass)
3         Err.Clear: On Error GoTo NewControl_Error
          '产生动态控件
4         If objPart Is Nothing Then
5             Set objCrl = objParent.Controls.Add(strControlClass, strName)
6         Else
7             Set objCrl = objParent.Controls.Add(strControlClass, strName)
8             Set objCrl.Container = objPart
9             objCrl.Move 0, 0, objPart.Width, objPart.Height
10            objCrl.ZOrder
11            objCrl.Visible = True
12        End If
13        If strControlClass = "zlLisControl.ucLisIDKind" Then
14            If Not objCrl.object.InitControl(objParent, gcnLisOracle, gUserInfo.DBUser) Then
15                Exit Function
16            End If
17        End If
          
18        Set NewControl = objCrl
          


19        Exit Function
NewControl_Error:
20        Call WriteErrLog("zlPublicHisCommLis", "mdlAboutReport", "执行(NewControl)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
21        Err.Clear

End Function

Public Function PtintOldReport(objFrm As Object, lngSampleID As Long, Optional lngPaintID As Long, Optional byRunMode As Byte = 2, Optional ByVal intSpecialPrintPage As Integer, Optional strErr As String) As Boolean
  '打印老版报告
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim lng医嘱ID As Long, lng发送号 As Long
    Dim strSQL As String
    Dim strChart(0 To 8) As String

    On Error GoTo PtintOldReport_Error

    strSQL = "select 发送号, a.医嘱id from 病人医嘱发送 a , 病人医嘱记录 b,检验标本记录  c where b.id = a.医嘱id and  a.医嘱id =c.医嘱id  and c.id = [1]"
    Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "报告打印", lngSampleID)
    If rsTmp.EOF = False Then
        lng发送号 = Val("" & rsTmp("发送号"))
        lng医嘱ID = Val("" & rsTmp("医嘱id"))
    End If

    If GetReportCode(lng医嘱ID, lng发送号, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
        If byRunMode = 3 Then
            FunReportPrintSetHis gcnHisOracle, 100, strReportCode, objFrm
        Else
            If ReadSampleImage(lngSampleID, strChart, strErr, 10) = False Then
                Exit Function
            End If
            Call FunReportOpenHis(gcnHisOracle, 100, strReportCode, objFrm, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, "医嘱ID=" & lng医嘱ID, _
                                "病人ID=" & lngPaintID, "标本ID=" & lngSampleID, _
                                "图形1=" & strChart(0), "图形2=" & strChart(1), "图形3=" & strChart(2), "图形4=" & strChart(3), _
                                "图形5=" & strChart(4), "图形6=" & strChart(5), "图形7=" & strChart(6), "图形8=" & strChart(7), _
                                "图形9=" & strChart(8), "DisabledPrint=1", intSpecialPrintPage, byRunMode)
        End If
    End If
    PtintOldReport = True


    Exit Function
PtintOldReport_Error:
    Call WriteErrLog("zlPublicHisCommLis", "mdlAboutReport", "执行(PtintOldReport)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
    Err.Clear

End Function

Public Function PrintNewReport(objFrm As Object, lngSampleID As Long, Optional byRunMode As Byte = 2, Optional ByVal blnDoctorShow As Boolean, Optional ByVal strPrivs As String, Optional ByVal intSpecialPrintPage As Integer, Optional strErr As String) As Boolean
'功能       打印报告
    Dim intCount As Integer
    Dim strNO As String
    Dim intSel As Integer
    Dim strChart(0 To 8) As String
    Dim strSQL As String
    Dim strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim rsReportFormat As ADODB.Recordset
    Dim lngPrintCount As Long

    On Error GoTo PrintNewReport_Error

    strSQL = "select b.id 仪器id ,b.名称 仪器名称,b.仪器类别,Nvl(a.病人来源,1) 病人来源,a.报告时间,a.阳性报告,a.标本序号,a.医生站打印,审核人  from 检验报告记录 a,检验仪器记录 b where a.仪器id = b.id and a.id = [1]"
    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "报告打印", lngSampleID)

    If rsTmp.RecordCount = 0 Then Exit Function
    
    
    '对比打印次数和参数
    If blnDoctorShow Then
        lngPrintCount = Val(ComGetPara(Sel_Lis_DB, "医生工作站报告打印次数", 2500, 2500, 1))
        If lngPrintCount > 0 Then
            If Val(rsTmp("医生站打印") & "") >= lngPrintCount And Val(rsTmp("病人来源") & "") = 2 Then
                strErr = "超出打印次数禁止打印"
                PrintNewReport = False
                Exit Function
            End If
        End If

    Else
        If rsTmp("审核人") & "" = "" And byRunMode = 2 Then
            If InStr(";" & strPrivs & ";", ";未审核报告打印;") Then
                strErr = "未审核报告不能打印"
                Exit Function
            End If
        End If
    End If
    
    strSQL = "select id,编码,名称,门诊单据,住院单据,体检单据,院外单据,门诊格式,住院格式,体检格式,院外格式,格式数量," & vbNewLine & _
           "       'ZLLISBILL' || Trim(To_Char(门诊单据, '00000')) || '-2' 门诊单据号," & vbNewLine & _
           "       'ZLLISBILL' || Trim(To_Char(住院单据, '00000')) || '-2' 住院单据号," & vbNewLine & _
           "       'ZLLISBILL' || Trim(To_Char(体检单据, '00000')) || '-2' 体检单据号," & vbNewLine & _
           "       'ZLLISBILL' || Trim(To_Char(院外单据, '00000')) || '-2' 院外单据号," & vbNewLine & _
           "       'ZLLISBILL' || Trim(To_Char(门诊格式, '00000')) || '-2' 门诊格式号," & vbNewLine & _
           "       'ZLLISBILL' || Trim(To_Char(住院格式, '00000')) || '-2' 住院格式号," & vbNewLine & _
           "       'ZLLISBILL' || Trim(To_Char(体检格式, '00000')) || '-2' 体检格式号," & vbNewLine & _
           "       'ZLLISBILL' || Trim(To_Char(院外格式, '00000')) || '-2' 院外格式号" & vbNewLine & _
             "from 检验仪器记录 where id = [1] "

    Set rsReportFormat = ComOpenSQL(Sel_Lis_DB, strSQL, "检验技师站", Val(rsTmp("仪器ID") & ""))


    rsReportFormat.Filter = "id=" & Val(rsTmp("仪器ID") & "")
    If Val(rsTmp("仪器类别")) = 1 Then
        If Val(rsTmp("阳性报告") & "") = 1 Then
            '阳性
            intSel = 0
        Else
            '阴性
            intSel = 1
        End If
    Else
        intCount = GetSampleValCount(lngSampleID)
        '没有结果时提示
        If intCount = 0 Then
            Exit Function
        End If
        If rsReportFormat.RecordCount > 0 Then
            If Val(rsReportFormat("格式数量") & "") > 0 Then
                If intCount > Val(rsReportFormat("格式数量") & "") Then
                    intSel = 0
                Else
                    intSel = 1
                End If
            End If
        Else
            intSel = 0
        End If
    End If

    Select Case Val(rsTmp("病人来源"))
    Case 1
        If intSel = 0 Then
            strNO = rsReportFormat("门诊单据号")
        Else
            strNO = rsReportFormat("门诊格式号")
        End If
    Case 2
        If intSel = 0 Then
            strNO = rsReportFormat("住院单据号")
        Else
            strNO = rsReportFormat("住院格式号")
        End If
    Case 3
        If intSel = 0 Then
            strNO = rsReportFormat("住院单据号")
        Else
            strNO = rsReportFormat("住院格式号")
        End If
    Case 4
        If intSel = 0 Then
            strNO = rsReportFormat("院外单据号")
        Else
            strNO = rsReportFormat("院外格式号")
        End If
    Case Else
        If intSel = 0 Then
            strNO = rsReportFormat("门诊单据号")
        Else
            strNO = rsReportFormat("门诊格式号")
        End If
    End Select

    If byRunMode = 3 Then
        If strNO <> "" Then
            FunReportPrintSet gcnLisOracle, gSysInfo.SysNo, strNO, objFrm
        End If
    Else
        '读图像
        strTmp = "开始读入图像:" & Now & vbCrLf
        If ReadSampleImage(lngSampleID, strChart, strErr, 25) = False Then
            Exit Function
        End If
        strTmp = strTmp & "读入图像完成:" & Now & vbCrLf

        FunReportOpen gcnLisOracle, gSysInfo.SysNo, strNO, objFrm, "标本ID=" & lngSampleID, "图形1=" & strChart(0), "图形2=" & strChart(1), "图形3=" & strChart(2), _
                      "图形4=" & strChart(3), "图形5=" & strChart(4), "图形6=" & strChart(5), "图形7=" & strChart(6), "图形8=" & strChart(7), _
                      "图形9=" & strChart(8), "DisabledPrint=1", intSpecialPrintPage, byRunMode
        strTmp = strTmp & "打印完成:" & Now & vbCrLf

        '对于审核过的标本标识
        strSQL = "Zl_检验报告打印_Edit(1," & lngSampleID & ",1)"
        Call ComExecuteProc(Sel_Lis_DB, strSQL, "打印标本")
        strTmp = strTmp & "完成打印:" & Now

        SaveDBLog 18, 6, lngSampleID, "打印", "报告打印", 2500, "临床实验室管理"
    End If

    PrintNewReport = True

    '发送刷新科内概况已打印标签申请
    Call SendMessage("RefreshDeptSurvey7")


    Exit Function
PrintNewReport_Error:
    Call WriteErrLog("zlPublicHisCommLis", "mdlAboutReport", "执行(PrintNewReport)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
    Err.Clear

End Function

