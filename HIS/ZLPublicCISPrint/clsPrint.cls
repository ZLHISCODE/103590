VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const gstrSysName As String = "中联软件"

Public mobjRegister As Object               '注册授权部件zlRegister
Private mobjEMR  As Object
Private mobjRichEMR As Object
Private mobjPublicPACS As Object

Private mstrPrintDocIDs As String
Private mstrUser As String
Private mstrTipInfo As String

Private WithEvents mclsDockAduits   As zlRichEPR.clsDockAduits
Attribute mclsDockAduits.VB_VarHelpID = -1
'首页信息;医嘱记录;住院病历;护理记录;护理病历;诊疗报告;疾病证明;知情文件;临床路径

Private Function UpdateRS(ByVal rsMed As ADODB.Recordset, ByVal strXML As String) As ADODB.Recordset
'功能:自定义记录集,用来封装阶段信息
'参数:
'strXML格式如下
'<items>
'    <item>
'        <id></id>
'        <file_path></file_path>
'    </item>
'</items>
    Dim xmlDoc As New DOMDocument
    Dim xNode As IXMLDOMNode
    Dim xNodeList As IXMLDOMNodeList
    On Error GoTo errH
 

    '读取网关响应数据（XML格式）
    xmlDoc.loadXML (strXML)
    Set xNodeList = xmlDoc.selectNodes(".//item")
    For Each xNode In xNodeList
        rsMed.Filter = "ID='" & xNode.selectSingleNode(".//id").Text & "'"
        If Not rsMed.EOF Then
            On Error Resume Next
            rsMed!文件路径 = xNode.selectSingleNode(".//file_path").Text
            On Error GoTo 0
            rsMed!标记 = 1
        End If
    Next
    rsMed.Filter = "标记=1"
    Set UpdateRS = rsMed
    Exit Function
errH:
  MsgBox Err.Description & vbCrLf & "MakeRS" & "行 " & Erl(), vbExclamation + vbOKOnly, gstrSysName
    
End Function

Private Function InitEMR(ByVal strUser As String, ByVal strPwd As String, ByRef strError As String) As Boolean
    Call UpdateEmrInterface(strUser, strPwd)
    If Not mobjEMR Is Nothing Then
        If Not mobjEMR.IsInited Or mobjEMR.IsOffline Then
            Set mobjEMR = Nothing
        Else
            On Error Resume Next
            Set mobjRichEMR = CreateObject("zlRichEMR.clsDockContent")
            If Err <> 0 Then
                strError = "新版病历组件创建失败，请联系管理员检查是否正确安装!"
                Set mobjEMR = Nothing
            End If
            Err.Clear
            If Not mobjRichEMR Is Nothing Then Call mobjRichEMR.Init(mobjEMR, gcnOracle, glngSys, 0)
        End If
    End If
    InitEMR = (Not mobjEMR Is Nothing)
End Function

Public Function InitPrint(ByVal strServer As String, ByVal strUser As String, ByVal strPass As String, ByVal strDBPwd As String, Optional ByRef strError As String) As Boolean
'功能:初始化连接对象、公用对象
'参数:
'strServer-服务名
'strUser-用户名
'strPass-加密密码(导航台录入)
'strDBPwd-数据库密码(新版病历需要)
'返回参数:
'strError-错误提示信息
    If mobjRegister Is Nothing Then
        Set mobjRegister = GetObject("", "zlRegister.clsRegister")
        Err.Clear
        If mobjRegister Is Nothing Then
            Set mobjRegister = CreateObject("zlRegister.clsRegister")
            Err.Clear
            If mobjRegister Is Nothing Then
                MsgBox "创建zlRegister部件对象失败,请检查文件是否存在并且正确注册。", vbExclamation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    On Error GoTo errHand
    glngSys = 100
    mstrUser = strUser
    If gcnOracle.State = adStateOpen Then gcnOracle.Close
    Set gcnOracle = mobjRegister.GetConnection(strServer, strUser, strPass, True, , strError)

    If strError <> "" Then MsgBox strError, vbInformation
    If gcnOracle.State = adStateClosed Then Exit Function
    Set gobjComlib = CreateObject("zl9ComLib.clsComLib")
    gobjComlib.InitCommon gcnOracle
    Set gobjDatabase = CreateObject("zl9ComLib.clsDatabase")
    Set gobjCommFun = CreateObject("zl9ComLib.clsCommFun")
    Set gobjReport = CreateObject("zl9Report.clsReport")
    gobjReport.InitOracle gcnOracle
    Err.Clear
    Call gobjCommFun.PDFInitialize(strError)
    If Err.Number <> 0 Or strError <> "" Then
        MsgBox "PDF设备初始化失败:" & strError, vbExclamation, gstrSysName: Exit Function
    End If
    
    If mobjPublicPACS Is Nothing Then
        On Error Resume Next
        Set mobjPublicPACS = CreateObject("zlPublicPACS.clsPublicPACS")
        Err.Clear: On Error GoTo 0
        If Not mobjPublicPACS Is Nothing Then
            Call mobjPublicPACS.InitInterface(gcnOracle, strUser)
        End If
        If mobjPublicPACS Is Nothing Then
            strError = "PACS公共部件未创建成功！"
        End If
    End If
    Call InitEMR(strUser, strDBPwd, strError)
    InitPrint = True
    Exit Function
errHand:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function PrintDocument(ByVal lngPatiID As Long, ByVal lngVisitID As Long, ByVal strFilePath As String, Optional ByVal strXML As String, Optional ByVal blnMerge As Boolean, Optional ByRef strNoPDF As String) As Boolean
'功能:输出指定病人全部或指定文档
'   lngPatiID-病人ID
'   lngVisitID-主页ID
'   strFilePath-PDF文件输出路径()
'   strXML-用户指定输出的文件清单,其形式为XML类型的字符串
'   一份文档对应一个<Item>节点 示例如下:
'   <items>
'       <item>
'           <id></id>           --可通过 GetPrintList获得id
'           <file_path></file_path>  --该节点指定输出文件的绝对路径(如:D:\TEST\首页.PDF)；没有该节点时按默认方式生成文件名。
'       </item>
'   </items>
'   blnMerge- T-将输出文档合并为一个文档; F-不合并文档
'   strNoPDF-返回未输出成功的PDF文档名
'
    Dim rsPati As ADODB.Recordset
    Dim lngDeptID As Long
    Dim objKernel As Object
    Dim strDeviceName As String
    Dim strPrint As String
    Dim strFileName As String
    Dim strPath As String
    Dim strType As String
    Dim blnDataMove As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim objRichEPR As Object
    Dim objTendsNew As Object
    Dim objPath As Object
    Dim objDockPath As Object
    Dim objFSO As New Scripting.FileSystemObject        'FSO对象
    Dim strReportName As String
    Dim strPDFTK As String
    Dim varParam As Variant
    Dim blnPrint As Boolean
    Dim lngFileCount As Long
    
    On Error GoTo errHand
    mstrPrintDocIDs = ""
    Set rsPati = GetPatiInfo(lngPatiID, lngVisitID)
    If rsPati.EOF Then Exit Function
    lngDeptID = Val(rsPati!出院科室ID & ""): blnDataMove = NVL(rsPati!数据转出, 0) = 1
    If Not objFSO.FolderExists(strFilePath) Then
        Call objFSO.CreateFolder(strFilePath)
    End If
    '检测是否存在TinyPDF(32位系统) Foxit Reader PDF Printer (64位系统)打印机
    strDeviceName = gobjCommFun.PDFPrinterDeviceName()
    
'    '病历文件
    Set rsTmp = GetCISStruct(lngPatiID, lngVisitID, "", blnDataMove)
    If strXML <> "" Then
        Set rsTmp = UpdateRS(rsTmp, strXML)
    End If
    strPath = strFilePath & "\" & NVL(rsPati!姓名) & "_" & lngPatiID & "_" & lngVisitID & "_"
    Do While Not rsTmp.EOF
        varParam = Split(rsTmp!参数 & "", ";")
        strFileName = rsTmp!文件路径 & ""
        blnPrint = False
        If rsTmp!ID Like "R11K*" Then
            If strReportName = "" Then strReportName = GetMedRecReportName(lngDeptID)
            If strFileName = "" Then strFileName = strPath & rsTmp!名称 & ".PDF"
            mstrTipInfo = "开始输出：" & strFileName
            Call SaveSetting("ZLSOFT", "私有模块\" & mstrUser & "\zl9Report\LocalSet\" & strReportName, "Printer", strDeviceName)
            blnPrint = gobjReport.ReportOpen(gcnOracle, glngSys, strReportName, Nothing, "病人ID=" & lngPatiID, "主页ID=" & lngVisitID, "ReportFormat=" & varParam(0), "PDF=" & strFileName, 4)
        ElseIf rsTmp!ID Like "R12K*" Then
            '医嘱清单
            If strFileName = "" Then strFileName = strPath & rsTmp!名称 & ".PDF"
            mstrTipInfo = "开始输出：" & strFileName
            If objKernel Is Nothing Then
                Set objKernel = CreateObject("zlCISKernel.clsCISKernel")
                Call objKernel.InitCISKernel(gcnOracle, Nothing, glngSys, "")
            End If
            blnPrint = objKernel.zlPrintAdvice(Nothing, lngPatiID, lngVisitID, 0, CInt(varParam(0)), strFileName, 4)
        ElseIf rsTmp!ID Like "R2K*" Or rsTmp!ID Like "R4K*" Or rsTmp!ID Like "R5K*" Or rsTmp!ID Like "R6K*" Then            '2-住院病历;4-护理病历;5-疾病证明;6-知情文件
            If objRichEPR Is Nothing Then
                Set objRichEPR = CreateObject("zlRichEPR.cRichEPR")
                Call objRichEPR.InitRichEPR(gcnOracle, Nothing, glngSys, False)
                Set mclsDockAduits = CreateObject("zlRichEPR.clsDockAduits")
            End If
            If InStr("," & mstrPrintDocIDs, "," & varParam(0) & ",") = 0 Then    '本次没打过     '
                If rsTmp!上级id = "R2" Then
                    strType = "住院病历"
                ElseIf rsTmp!上级id = "R4" Then
                    strType = "护理病历"
                ElseIf rsTmp!上级id = "R5" Then
                    strType = "疾病证明"
                ElseIf rsTmp!上级id = "R6" Then
                    strType = "知情文件"
                End If
                If strFileName = "" Then strFileName = strPath & strType & "_" & varParam(3) & ".PDF"
                mstrTipInfo = "开始输出：" & strFileName
                strPrint = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
                Call gobjCommFun.PDFFile(strFileName)
                blnPrint = mclsDockAduits.zlPrintDocument(3, 2, Val(varParam(0)), strPrint, blnDataMove)
            End If
        ElseIf rsTmp!ID Like "R3K*" Then          '护理记录
            If varParam(5) = "1" Then   '旧版'
                If objRichEPR Is Nothing Then
                    Set objRichEPR = CreateObject("zlRichEPR.cRichEPR")
                    Call objRichEPR.InitRichEPR(gcnOracle, Nothing, glngSys, False)
                    Set mclsDockAduits = CreateObject("zlRichEPR.clsDockAduits")
                End If
                If Val(varParam(1)) = -1 Then
                    '体温单
                    If strFileName = "" Then strFileName = strPath & "体温单_" & varParam(3) & ".PDF"
                    mstrTipInfo = "开始输出：" & strFileName
                    strPrint = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
                    Call gobjCommFun.PDFFile(strFileName)
                    Call mclsDockAduits.zlRefreshTendBody(lngPatiID, lngVisitID, Val(varParam(0)), Val(varParam(4)), blnDataMove)
                    blnPrint = mclsDockAduits.zlPrintDocument(1, 2, , strPrint, blnDataMove)
                Else
                    '护理记录
                    If strFileName = "" Then strFileName = strPath & "护理记录_" & varParam(3) & ".PDF"
                    mstrTipInfo = "开始输出：" & strFileName
                    strPrint = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
                    Call gobjCommFun.PDFFile(strFileName)
                    Call mclsDockAduits.zlRefresh(3, Val(varParam(3)), lngPatiID, lngVisitID, Val(varParam(0)), CStr(varParam(2)), , Val(varParam(4)), blnDataMove)
                    blnPrint = mclsDockAduits.zlPrintDocument(2, 2, , strPrint, blnDataMove)
                End If
            Else  '新
                If strFileName = "" Then strFileName = strPath & Decode(Val(varParam(1)), -1, "体温单", 1, "产程图", "护理记录") & "_" & Val(varParam(3)) & ".PDF"
                mstrTipInfo = "开始输出：" & strFileName
                Call gobjCommFun.PDFFile(strFileName)
                If objTendsNew Is Nothing Then
                    Set objTendsNew = CreateObject("zl9TendFile.clsTendFile")
                    Call objTendsNew.InitTendFile(gcnOracle, glngSys)
                End If
                blnPrint = objTendsNew.zlPrintDocument(lngPatiID, lngVisitID, Val(varParam(4)), Val(varParam(0)), Val(varParam(3)), Decode(Val(varParam(1)), -1, 1, 1, 3, 2), strDeviceName, True)
            End If
        ElseIf rsTmp!ID Like "R7K*" Then                 '诊疗报告
            '系统号,FileName,ID
            If objRichEPR Is Nothing Then
                Set objRichEPR = CreateObject("zlRichEPR.cRichEPR")
                Call objRichEPR.InitRichEPR(gcnOracle, Nothing, glngSys, False)
                Set mclsDockAduits = CreateObject("zlRichEPR.clsDockAduits")
            End If
            If strFileName = "" Then strFileName = strPath & "诊疗报告_" & varParam(3) & "_" & Val(varParam(0)) & ".PDF"
            mstrTipInfo = "开始输出：" & strFileName
            Call gobjCommFun.PDFFile(strFileName)
            strPrint = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
            blnPrint = mclsDockAduits.zlPrintDocument(4, 2, Val(varParam(0)), strPrint, blnDataMove)
        ElseIf rsTmp!ID Like "R7P*" Then
            If Not mobjPublicPACS Is Nothing Then
                If strFileName = "" Then strFileName = strPath & "诊疗报告_" & rsTmp!名称 & ".PDF"
                mstrTipInfo = "开始输出：" & strFileName
                blnPrint = mobjPublicPACS.PrintReport(varParam(0), strFileName, False)      'True预览
            End If
        ElseIf rsTmp!ID Like "R8K*" Then
            If strFileName = "" Then strFileName = strPath & rsTmp!名称 & ".PDF"
            mstrTipInfo = "开始输出：" & strFileName
            If objDockPath Is Nothing Then
                Set objDockPath = CreateObject("zlCISPath.clsDockPath")
                Set objPath = CreateObject("zlCISPath.clsCISPath")
                Call objPath.InitCISPath(gcnOracle, Nothing, glngSys, "")
            End If
            blnPrint = objDockPath.zlFuncPathTableOutPut(4, True, strFileName, lngPatiID, lngVisitID, strDeviceName)
        ElseIf InStr(rsTmp!ID, "R") = 0 And InStr(rsTmp!参数, ";EMR;") > 0 Then
            '存在共享文档时只打印输出一次
            If InStr("," & mstrPrintDocIDs, "," & varParam(0) & ",") = 0 And Not mobjRichEMR Is Nothing Then
                If varParam(1) <> "" Then
                    Call mobjRichEMR.zlShowDoc(varParam(0), varParam(1))
                Else
                    Call mobjRichEMR.zlShowDoc(varParam(0), "")
                End If
                If strFileName = "" Then strFileName = strPath & varParam(2) & varParam(0) & ".PDF"
                mstrTipInfo = "开始输出：" & strFileName
                
                Call gobjCommFun.PDFFile(strFileName)
                strPrint = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
                 
                Err.Number = 0: On Error Resume Next
                blnPrint = mobjRichEMR.zlPrintDoc(False, strPrint)
                If Err.Number = 450 Then
                    Err.Number = 0
                    blnPrint = mobjRichEMR.zlPrintDoc(False)
                End If
                Err.Clear: On Error GoTo 0
                mstrPrintDocIDs = mstrPrintDocIDs & varParam(0) & ","  '
            End If
        End If
        Call gobjCommFun.PDFFileSuccess
        
        If Dir(strFileName) <> "" Then
            If blnMerge Then
                strPDFTK = strPDFTK & strFileName & Space(1)
                lngFileCount = lngFileCount + 1
            End If
        Else
            strNoPDF = strNoPDF & strFileName & ";"
        End If
        mstrTipInfo = ""
        rsTmp.MoveNext
    Loop
    
    If strPDFTK <> "" And lngFileCount > 1 Then
        Shell "pdftk.exe " & strPDFTK & " cat output " & strPath & "病案合并.PDF"
    End If
    mstrTipInfo = ""
    PrintDocument = True
    
    Exit Function
errHand:
    mstrTipInfo = ""
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetPrintList(ByVal lngPatiID As Long, ByVal lngVisitID As Long) As String
'功能:获取指定病人的可输出文件清单。
'返回值:字符串 格式如下
'   <items>
'       <item>
'           <id></id>--文件ID
'           <type></type> --住院病历\护理记录\护理病历\疾病证明\知情文件\诊疗报告\首页信息\医嘱记录\临床路径
'           <file_name></file_name> --文档名称
'       </item>
'   </items>

    Dim rsPati As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strXML As String
    
    Set rsPati = GetPatiInfo(lngPatiID, lngVisitID)
    If rsPati.EOF Then Exit Function
    Set rsTmp = GetCISStruct(lngPatiID, lngVisitID, "", NVL(rsPati!数据转出, 0) = 1)
    If rsTmp Is Nothing Then Exit Function
    For i = 1 To rsTmp.RecordCount
        strXML = strXML & "<item>"
        strXML = strXML & "<id>" & rsTmp!ID & "</id>"
        strXML = strXML & "<type>" & GetFileType(rsTmp!上级id) & "</type>"
        strXML = strXML & "<file_name>" & rsTmp!名称 & "</file_name>"
        strXML = strXML & "</item>"
        rsTmp.MoveNext
    Next
    If strXML <> "" Then strXML = "<items>" & strXML & "</items>"
    GetPrintList = strXML
End Function

Private Function GetFileType(ByVal strType As String) As String
    ''1-门诊病历;2-住院病历;3-护理记录;4-护理病历;5-疾病证明;6-知情文件;7-诊疗报告,11-首页信息,12-医嘱记录,8-临床路径
    GetFileType = Split(",门诊病历,住院病历,护理记录,护理病历,疾病证明,知情文件,诊疗报告,临床路径,,,首页信息,医嘱记录", ",")(CInt(Replace(strType, "R", "")))
End Function


Private Sub UpdateEmrInterface(ByVal strUser As String, ByVal strPwd As String)
    If Not mobjEMR Is Nothing Then Exit Sub
    On Error Resume Next
    Err.Clear
    Set mobjEMR = CreateObject("zl9EmrInterface.ClsEmrInterface")
    If Err.Number = 0 Then
        Call mobjEMR.CheckUpdate1(strUser, "[DBPASSWORD]" & strPwd, False)
        If Err.Number <> 0 Then
            Err.Clear
            Call mobjEMR.CheckUpdate(strUser, strPwd)
        End If
    Else
        Set mobjEMR = Nothing
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
End Sub

Private Function GetCISStruct(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str挂号单 As String, ByVal blnDataMove As Boolean) As ADODB.Recordset
'参数:lng主页ID 住院病人为主页id,门诊病人为挂号id
    Dim strSQL As String, strSQL1 As String
    Dim rsTmp As ADODB.Recordset
    Dim rsMedRec As ADODB.Recordset
    Dim strRptIDs As String
    
    On Error GoTo errH
    '1-门诊病历;2-住院病历;3-护理记录;4-护理病历;5-疾病证明;6-知情文件;7-诊疗报告,11-首页信息,12-医嘱记录,8-临床路径;9-住院证;10-其他报表
    strSQL = " Select 'R11K1' As ID, 'R11' As 上级id, '首页正面' As 名称, '1' As 参数, '1' As 排序 From Dual" & _
            " Union All Select 'R11K2' As ID, 'R11' As 上级id, '首页反面' As 名称, '2' As 参数,'2' As 排序 From Dual"
    If Val(gobjDatabase.GetPara("病案首页标准", glngSys, p住院医生站, "0")) = 2 Then
        strSQL = strSQL & _
              " Union All Select 'R11K3' As ID, 'R11' As 上级id, '首页附页一' As 名称, '3' As 参数, '3' As 排序 From Dual" & _
              " Union All Select 'R11K4' As ID, 'R11' As 上级id, '首页附页二' As 名称, '4' As 参数, '4' As 排序 From Dual"
    End If
    
    '医嘱部分
    strSQL = strSQL & " Union All Select 'R12K1' As ID, 'R12' As 上级id, '临时医嘱' As 名称, '1' As 参数,'1' As 排序 From Dual"
    strSQL = strSQL & " Union All Select 'R12K2' As ID, 'R12' As 上级id, '长期医嘱' As 名称, '0' As 参数,'2' As 排序 From Dual"
    
    '临床路径
    strSQL1 = "Select 1 From 病人临床路径 Where 病人id = [1] And 主页id = [2] And 状态 <> 0"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL1, "检查是否存有效的临床路径数据", lng病人ID, lng主页ID)
    If rsTmp.RecordCount > 0 Then
        strSQL = strSQL & _
          " Union All Select 'R8K' As ID, 'R8' As 上级id, '临床路径' As 名称, '' As 参数, '1' As 排序 From Dual"
    End If
    
    '病历部分
    'ID=上级ID+K病历ID
    strSQL = strSQL & " Union All Select A.上级id||'K'||Trim(To_Char(A.ID)) As ID,A.上级id," & _
        "       Decode(A.医嘱id,Null,A.名称||'('||To_Char(A.创建时间, 'YYYY-MM-DD')||')',A.名称||'：'||B.医嘱内容||'('||To_Char(A.创建时间, 'YYYY-MM-DD')||')') As 名称," & _
        "       Trim(To_Char(A.ID))||';'||Decode(A.医嘱id,Null,'0',Trim(To_Char(A.医嘱id))) || ';'|| A.RISID||';'|| A.名称||';'||A.编辑方式 ||';'||C.操作类型 As 参数,排序 " & _
        " From (Select A.ID, 'R'||A.病历种类 As 上级id, A.病历名称 As 名称,C.医嘱id,C.RISID,A.病历种类,A.编辑方式,A.创建时间,To_Char(A.创建时间,'YYYY-MM-DD HH24:MI:SS') As 排序" & _
        "       From 电子病历记录 A,病人医嘱报告 C " & _
        "       Where A.病人id = [1] And A.主页id = [2] And (A.病人来源=2 And [3]=1 Or Nvl(A.病人来源,0)<>2 And [3]=0)" & _
        "           And C.病历id(+)=A.ID And A.病历种类 In (1, 2, 3, 4, 5, 6, 7)" & _
        "       ) A,病人医嘱记录 B,诊疗项目目录 C Where A.医嘱id=B.Id(+) And B.诊疗项目ID = C.ID(+)"
    '护理部分
    'ID=上级ID+K文件ID,0,科室ID
    '参数=科室ID;保留;开始～截止;文件ID
    '检查本次病人是使用的是老板还是新版
    strSQL1 = "Select 1 From 病人护理记录 A Where a.病人id = [1] And a.主页id = [2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL1, "检查是否存在老板数据", lng病人ID, lng主页ID)
    If rsTmp.RecordCount > 0 Then
        strSQL = strSQL & " Union All" & _
            " Select 'R3K'||Trim(To_Char(A.ID)) As ID,'R3' As 上级id," & _
            "       A.名称||'('||B.名称||'：'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI') || '～' ||To_Char(A.截止, 'YYYY-MM-DD HH24:MI') || ')' As 名称," & _
            "       Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(保留,0)))||';'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI')||'～'||To_Char(A.截止, 'YYYY-MM-DD HH24:MI')||';'||Trim(To_Char(A.ID)) ||';'||'0;1' As 参数,To_Char(a.开始,'YYYY-MM-DD HH24:MI:SS') As 排序" & _
            " From (" & _
            "   Select F.ID, F.编号, F.名称, R.开始, R.截止, R.科室id, 保留" & _
            "   From (" & _
            "       Select ID, 编号, 名称, 3 As 护理级别, 通用, 0 As 科室id, 保留" & _
            "          From 病历文件列表 Where 种类 = 3 And 保留 < 0" & _
            "       Union All" & _
            "       Select L.ID, L.编号, L.名称, F.报表 As 护理级别, L.通用, A.科室id, L.保留" & _
            "          From 病历页面格式 F, 病历文件列表 L, 病历应用科室 A" & _
            "          Where L.种类 = 3 And L.保留 = 0 And L.种类 = F.种类 And L.编号 = F.编号 And L.ID = A.文件id(+)" & _
            "       ) F,(" & _
            "       Select R.科室id, Nvl(Min(R.护理级别), 3) As 护理级别, Min(R.发生时间) As 开始, Max(R.发生时间) As 截止" & _
            "          From 病人护理记录 R" & _
            "          Where R.病人来源 = 2 And R.病人id = [1] And Nvl(R.主页id, 0) = [2] And Nvl(R.婴儿, 0) = 0" & _
            "          Group By R.科室id" & _
            "       ) R" & _
            "       Where (F.通用 = 1 Or F.通用 = 2 And F.科室id = R.科室id) And F.护理级别 >= R.护理级别" & _
            "   ) A, 部门表 B Where A.科室id = B.ID "
    Else
        strSQL = strSQL & " Union All" & _
                " Select 'R3K'||Trim(To_Char(A.ID)) As ID,'R3' As 上级id," & vbNewLine & _
                "     A.名称||'('||B.名称||'：'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI') || '～' ||To_Char(A.截止, 'YYYY-MM-DD HH24:MI') || ')' As 名称," & vbNewLine & _
                "      Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(保留,0)))||';'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI')||'～'||To_Char(A.截止, 'YYYY-MM-DD HH24:MI')||';'||Trim(To_Char(A.ID))||';'||Trim(To_Char(A.婴儿))||';2' As 参数,To_Char(a.开始,'YYYY-MM-DD HH24:MI:SS') As 排序" & vbNewLine & _
                " From (" & vbNewLine & _
                "   Select R.ID, F.编号, R.名称,R.婴儿, R.开始, NVL(R.截止,nvl(R.时间,R.开始)) 截止, R.科室id, 保留" & vbNewLine & _
                "   From (" & vbNewLine & _
                "       Select L.ID, L.编号, L.名称, F.报表 As 护理级别, L.通用, L.保留" & vbNewLine & _
                "          From 病历页面格式 F, 病历文件列表 L" & vbNewLine & _
                "          Where L.种类 = 3 And L.种类 = F.种类 And L.编号 = F.编号 And (L.通用=1 OR L.通用=2)" & vbNewLine & _
                "" & vbNewLine & _
                "       ) F,(" & vbNewLine & _
                "       Select R.ID,R.科室id,R.文件名称 名称,R.格式ID,nvl(R.婴儿,0) 婴儿,Min(R.开始时间) As 开始, Max(R.结束时间) As 截止,MAX(T.发生时间) 时间" & vbNewLine & _
                "          From 病人护理文件 R,病人护理数据 T" & vbNewLine & _
                "          Where R.ID=T.文件ID(+) And R.病人id = [1] And Nvl(R.主页id, 0) = [2]" & vbNewLine & _
                "          Group By R.ID,R.文件名称,R.科室id,R.格式ID,R.婴儿" & vbNewLine & _
                "       ) R" & vbNewLine & _
                "       Where F.ID=R.格式ID" & vbNewLine & _
                "   ) A, 部门表 B Where A.科室id = B.ID And DECODE(A.保留,-1,0,A.婴儿)=A.婴儿"
    End If
    
    strSQL = " Select * From (" & strSQL & ") Order By Decode(上级id,Null,' ',上级id),排序"
    
    If blnDataMove And lng病人ID <> 0 Then
        strSQL = Replace(strSQL, "电子病历记录", "H电子病历记录")
        strSQL = Replace(strSQL, "病人护理记录", "H病人护理记录")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱报告", "H病人医嘱报告")
        strSQL = Replace(strSQL, "病人护理文件", "H病人护理文件")
        strSQL = Replace(strSQL, "病人护理数据", "H病人护理数据")
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "GetCISStruct", lng病人ID, lng主页ID, 1)
    Set rsMedRec = gobjDatabase.CopyNewRec(rsTmp, False, "", Array("文件路径", adVarChar, 500, Empty, "标记", adInteger, 1, Empty))
    'BUG:138397多条检验医嘱对应一份报告的时候去掉重复项目(按报告ID去重)
    strRptIDs = ""
    rsMedRec.Filter = "上级ID='R7'"
    Do While Not rsMedRec.EOF
        If InStr("," & strRptIDs & ",", "," & Split(NVL(rsMedRec!参数), ";")(0) & ",") = 0 Then
            strRptIDs = strRptIDs & "," & Split(NVL(rsMedRec!参数), ";")(0)
        Else
            rsMedRec.Delete
            If rsMedRec.EOF Then Exit Do
        End If
        rsMedRec.MoveNext
    Loop
    rsMedRec.Filter = ""
    'EMR
    Set rsTmp = GetEmrCISStruct(lng病人ID, lng主页ID)
    If Not rsTmp Is Nothing Then
        strRptIDs = ""
        Do While Not rsTmp.EOF
            If InStr("," & strRptIDs & ",", "," & rsTmp!ID & ",") = 0 Then
                rsMedRec.AddNew
                rsMedRec!ID = rsTmp!ID
                rsMedRec!上级id = rsTmp!上级id
                rsMedRec!名称 = rsTmp!名称
                rsMedRec!参数 = NVL(rsTmp!参数) & ";EMR;" & rsTmp!上级id
                rsMedRec.Update
                strRptIDs = strRptIDs & "," & rsTmp!ID
            End If
            rsTmp.MoveNext
        Loop
    End If
    '新版PACS
    Set rsTmp = Nothing
    If Not mobjPublicPACS Is Nothing Then
        Set rsTmp = mobjPublicPACS.zlDocGetList(lng病人ID, lng主页ID, str挂号单)
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                rsMedRec.AddNew
                rsMedRec!ID = "R7P" & rsTmp!报告ID
                rsMedRec!上级id = "R7"
                rsMedRec!名称 = rsTmp!文档标题 & ""
                rsMedRec!参数 = rsTmp!报告ID
                rsMedRec.Update
                rsTmp.MoveNext
            Loop
        End If
    End If
    rsMedRec.Filter = ""
    Set GetCISStruct = rsMedRec
    Exit Function
errH:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function GetEmrCISStruct(ByVal lngPatiID As Long, ByVal lngPageID As Long) As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strExtendTag As String, strReturn As String, strSQL As String, strSQLNew As String
    
    On Error GoTo errH
    If mobjEMR Is Nothing Then Set GetEmrCISStruct = Nothing: Exit Function
    strExtendTag = GetEMRIn_Tag(lngPatiID, lngPageID)
    If strExtendTag = "" Then Set GetEmrCISStruct = Nothing: Exit Function
    
    '上级ID，ID，名称，参数，图标
    strSQL = "Select Decode(e.Kind, '01', 'R1', '02', 'R2', '03', 'R4', '04', 'R5', '05', 'R6', 'R2') 上级id," & vbNewLine & _
            "       Nvl(d.Subdoc_Id, Rawtohex(d.Real_Doc_Id)) As ID, d.Subdoc_Id As 子文档id," & vbNewLine & _
            "       e.Title ||" & vbNewLine & _
            "        Decode(d.Completor, Null, ''," & vbNewLine & _
            "               '【 ' || d.Completor || ' 在' || To_Char(d.Complete_Time, 'yyyy-mm-dd hh24:mi') || '签名】') As 名称," & vbNewLine & _
            "       Rawtohex(d.Real_Doc_Id) || Decode(d.Subdoc_Id, Null, ';', ';' || d.Subdoc_Id) || ';' ||Nvl(d.Subdoc_Title, E.Title) As 参数, 'object_case' As 图标" & vbNewLine & _
            "From (Select Distinct d.Id, c.Antetype_Id, c.Subdoc_Id, c.Subdoc_Title, c.Real_Doc_Id, c.Complete_Time, c.Completor" & vbNewLine & _
            "       From Bz_Act_Log A, Bz_Act_Log D, Bz_Doc_Tasks C" & vbNewLine & _
            "       Where a.Extend_Tag = :etag And (a.Id = d.Id Or a.Id = d.Basiclog_Id) And d.Id = c.Actlog_Id And" & vbNewLine & _
            "             c.Real_Doc_Id Is Not Null) D, Antetype_List E" & vbNewLine & _
            "Where d.Antetype_Id = e.Id  And e.Title = Decode(e.Type, 3, d.Subdoc_Title, e.Title)" & vbNewLine & _
            "Order By Rawtohex(d.Real_Doc_Id), e.Code, d.Complete_Time"
            
    strSQLNew = "Select Decode(e.Kind, '01', 'R1', '02', 'R2', '03', 'R4', '04', 'R5', '05', 'R6', 'R2') 上级id," & vbNewLine & _
                "       Nvl(d.Subdoc_Id, Rawtohex(d.Real_Doc_Id)) As ID, d.Subdoc_Id As 子文档id," & vbNewLine & _
                "       e.Title ||" & vbNewLine & _
                "        Decode(d.Completor, Null, ''," & vbNewLine & _
                "               '【 ' || d.Completor || ' 在' || To_Char(d.Complete_Time, 'yyyy-mm-dd hh24:mi') || '签名】') As 名称," & vbNewLine & _
                "       Rawtohex(d.Real_Doc_Id) || Decode(d.Subdoc_Id, Null, ';', ';' || d.Subdoc_Id) || ';' ||Nvl(d.Subdoc_Title, E.Title) As 参数, 'object_case' As 图标" & vbNewLine & _
                "From (Select Distinct d.Id, c.Antetype_Id, c.Subdoc_Id, c.Subdoc_Title, c.Real_Doc_Id, c.Complete_Time, c.Completor, c.Order_No" & vbNewLine & _
                "       From Bz_Act_Log A, Bz_Act_Log D, Bz_Doc_Tasks C" & vbNewLine & _
                "       Where a.Extend_Tag = :etag And (a.Id = d.Id Or a.Id = d.Basiclog_Id) And d.Id = c.Actlog_Id And" & vbNewLine & _
                "             c.Real_Doc_Id Is Not Null And Nvl(c.Intead, 0) = 0) D, Antetype_List E" & vbNewLine & _
                "Where d.Antetype_Id = e.Id " & vbNewLine & _
                "Order By Rawtohex(d.Real_Doc_Id), e.Code, d.Order_No"
    
    Err.Clear
    On Error Resume Next
    strReturn = mobjEMR.OpenSQLRecordset(strSQLNew, strExtendTag & "^16^etag", rsTemp)
    If Err.Number <> 0 Or strReturn <> "" Then
        Err.Clear
        strReturn = mobjEMR.OpenSQLRecordset(strSQL, strExtendTag & "^16^etag", rsTemp)
    End If
    
    If strReturn <> "" Then
        MsgBox strReturn, vbCritical, gstrSysName
        Set GetEmrCISStruct = Nothing: Exit Function
    End If
    
    Set GetEmrCISStruct = rsTemp
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function GetEMRIn_Tag(ByVal lngPatiID As Long, ByVal lngPageID As Long) As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
 
    strSQL = "Select Nvl(a.Id, b.Id) ID" & vbNewLine & _
                "From (Select Max(ID) ID From 病人变动记录 Where 病人id = [1] And 主页id = [2] And 开始原因 = 2 And Nvl(附加床位, 0) = 0) A," & vbNewLine & _
                "     (Select Max(ID) ID From 病人变动记录 Where 病人id = [1] And 主页id = [2] And 开始原因 = 1 And Nvl(附加床位, 0) = 0) B"
    
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "读取病人入院ID", lngPatiID, lngPageID)
    
    If rsTmp Is Nothing Then Exit Function
    If NVL(rsTmp!ID) = "" Then Exit Function
    GetEMRIn_Tag = "BD_" & rsTmp!ID
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Sub Class_Terminate()
    glngSys = 0
    
    Set mobjRegister = Nothing
    Set gcnOracle = Nothing
    Set gobjReport = Nothing
    Set mobjPublicPACS = Nothing
    Set mobjRichEMR = Nothing
    Set mclsDockAduits = Nothing
    Set mobjEMR = Nothing
End Sub

Private Sub mclsDockAduits_AfterEprPrint(ByVal lngRecordId As Long)
    mstrPrintDocIDs = mstrPrintDocIDs & lngRecordId & ","
End Sub

Private Function GetMedRecReportName(ByVal lngDeptID As Long) As String
'功能:获取病案首页报表名称
    Dim strReportName As String
    
    Select Case Val(gobjDatabase.GetPara("病案首页标准", glngSys, p住院医生站, "0"))
         Case 0 '卫生部标准
             If gobjComlib.Sys.DeptHaveProperty(lngDeptID, "中医科") Then
                 strReportName = "ZL1_INSIDE_1261_4"
             Else
                 strReportName = "ZL1_INSIDE_1261_1"
             End If
         Case 1    '四川省标准
             If gobjComlib.Sys.DeptHaveProperty(lngDeptID, "中医科") Then
                 strReportName = "ZL1_INSIDE_1261_6"
             Else
                 strReportName = "ZL1_INSIDE_1261_5"
             End If
         Case 2    '云南省标准
             If gobjComlib.Sys.DeptHaveProperty(lngDeptID, "中医科") Then
                 strReportName = "ZL1_INSIDE_1261_8"
             Else
                 strReportName = "ZL1_INSIDE_1261_7"
             End If
         Case 3    '湖南省标准
             If gobjComlib.Sys.DeptHaveProperty(lngDeptID, "中医科") Then
                 strReportName = "ZL1_INSIDE_1261_10"
             Else
                 strReportName = "ZL1_INSIDE_1261_9"
             End If
     End Select
     GetMedRecReportName = strReportName
End Function
