Attribute VB_Name = "mdlDYEYMZYF"
Option Explicit

Public gobjSOAP As Object  '接口对象
Public gstrIP As String    '本机ip
Public gblnShowMsg As Boolean   '是否弹出对话框提示（自助收费需要）
Public gstrUnit As String   '用户注册用户名
Public gstrOutPut As String     '日志输出内容
Public gblnUpdateFlag As Boolean

Public Const GCST_UNIT_DYEY = "大连医科大学附属第二医院"
Public Const GCST_UNIT_YZSZYY = "扬州市中医院"
Public Const GCST_UNIT_JLSZXYY = "吉林市中心医院"
Public Const GCST_UNIT_CQFLQZYY = "重庆市涪陵区中医院"
Public Const GCST_UNIT_YNYXRMYY = "云南省玉溪市人民医院"
Public Const GCST_UNIT_YQMY = "阳泉煤业（集团）有限责任公司总医院"
Public Const GCST_UNIT_BTSZXYY = "包头市中心医院"

Public Const GINT_SEND_TYPE = 1           '0-仅开始发药流程，1-有开始发药，结束发药流程
Public Const GINT_STARTSEND_TYPE = 1      '0-按钮方式开始发药，1-刷卡方式开始发药
Public Const GBLN_OUTPUTLOG_DETAIL = True   '写日志时是否输出明细数据(上传到对方接口的明细数据)，如果是false只在出错时输出明细数据

'固定药房
Public Const GCST_DRUGID_DYEY = 176         '大连医科大学附属第二医院，门诊药房

Private Type IPINFO
    dwAddr As Long   ' IP address
    dwIndex As Long ' interface index
    dwMask As Long ' subnet mask
    dwBCastAddr As Long ' broadcast address
    dwReasmSize  As Long ' assembly size
    unused1 As Integer ' not currently used
    unused2 As Integer '; not currently used
End Type

Private Type MIB_IPADDRTABLE
    dEntrys As Long   'number of entries in the table
    mIPInfo(5) As IPINFO  'array of IP address entries
End Type
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Enum gType
    IntDrug = 101       '上传药品基础数据
    IntStore = 102      '上传药品库存数据
    IntDept = 104       '上传部门数据
    IntDetail = 201     '上传处方明细
    IntStartList = 202  '上传主处方单，开始发药
    IntEndList = 203    '上传主处方单，结束发药
    IntReturnAll = 205  '处方退费，全退模式
End Enum

Private mstrSQL As String

Private mobjFSO As New FileSystemObject

Public Function DYEY_MZ_TransData(ByVal intType As Integer, ByVal intOprId As Integer, ByVal strUserCode As String, _
    ByVal strUserName As String, ByVal arrXML As Variant, ByRef strReturn As String, Optional ByVal strNO As String, _
    Optional ByVal lngStockID As Long) As Boolean
'1.向WebService传递数据
'2.供接口函数调用
    Dim i As Integer
    Dim intRetval As Integer
    Dim strRETMSG As String
    Dim blnShow As Boolean
    Dim lngDrugStockID As Long
    Dim strOutput As String
    
    On Error GoTo errHandle
    If intType = gType.IntDrug Or intType = gType.IntStore Then
        If gblnShowMsg Then
            MsgBox "进入上传！", vbInformation, GSTR_MESSAGE
        Else
            strReturn = "进入上传！"
        End If
    End If
    If gstrIP = "" Then
        gstrIP = GetLocalIP
    End If
    
    For i = 0 To UBound(arrXML)
        If gobjSOAP.TransConsisData(intOprId, intType, CStr(arrXML(i)), gstrIP, strUserCode, strUserName, intRetval, strRETMSG) <> 1 Then
            If gblnShowMsg Then
                MsgBox strRETMSG, vbInformation + vbOKOnly, GSTR_MESSAGE
            Else
                strReturn = strRETMSG
            End If
            
            '日志输出
            strOutput = "SOAP.TransConsisData：" & vbNewLine & _
                        "   XML：" & arrXML(i) & vbNewLine & _
                        "   intRetval：" & intRetval & vbNewLine & _
                        "   strRETMSG：" & strRETMSG & vbNewLine
            Call OutputLog(strOutput)
            
            If blnShow Then frmDYEY_MZ_TransDrug.UnloadMe
            Exit Function
        End If
        
        If intType = gType.IntDrug Or intType = gType.IntStore Then
            If i = 0 Then
                frmDYEY_MZ_TransDrug.Show
                blnShow = True
            End If
            
            Call frmDYEY_MZ_TransDrug.ChangePrg(i + 1, UBound(arrXML) + 1, intType)
        ElseIf intType = gType.IntDetail Then
            '从上传信息中取药房ID
            lngDrugStockID = GetStockID(arrXML(i))
            
            If Not SetSendWin(lngDrugStockID, strNO, intRetval) Then
                If gblnShowMsg Then
                    MsgBox "调整处方的发药窗口失败！", vbCritical, GSTR_MESSAGE
                Else
                    strReturn = "调整处方的发药窗口失败！"
                End If
            End If
        End If
    Next
    
    DYEY_MZ_TransData = True
    If intType = gType.IntDrug Or intType = gType.IntStore Then
        If gblnShowMsg Then
            MsgBox "上传完成！", vbInformation, GSTR_MESSAGE
        Else
            strReturn = "上传完成！"
        End If
    End If
    Exit Function
errHandle:
    If gblnShowMsg Then
        If gintMode = 0 Then
            If gobjComLib.ErrCenter = 1 Then Resume
            Call gobjComLib.SaveErrLog
        Else
            MsgBox Err.Description, vbInformation, GSTR_SYSNAME
        End If
    End If
End Function

Public Function DYEY_MZ_TransData_CQFLQZYY(ByVal intType As Integer, ByVal intOprId As Integer, ByVal strUserCode As String, _
    ByVal strUserName As String, ByVal arrXML As Variant, ByRef strReturn As String, ByRef strOutput As String, _
    Optional ByVal strNO As String, Optional ByVal lngStockID As Long) As Boolean
    
'1.向WebService传递数据
'2.供接口函数调用
'3.适用接口：韦乐海茨CONSIS系统v4.3
    
    Dim i As Integer
    Dim blnShow As Boolean
    Dim lngDrugStockID As Long
    Dim strXML_In As String
    Dim strXML_Out As String
    Dim strOut_RETVAL As String
    Dim strOut_RETMSG As String
    Dim strOut_RETCODE As String
    Dim strTmp As String
    Dim strOutPutExeStep As String    '执行步骤，用于输出日志方便查找问题
    
    On Error GoTo errHandle
    
    strOutput = strOutput & vbCrLf & "调用函数：DYEY_MZ_TransData_CQFLQZYY"
    strOutput = strOutput & vbCrLf & "业务代码：" & intType
    
    If intType = gType.IntDrug Or intType = gType.IntStore Then
        If gblnShowMsg Then
            MsgBox "进入上传！", vbInformation, GSTR_MESSAGE
        Else
            strReturn = "进入上传！"
        End If
    End If
    
    If gstrIP = "" Then
        gstrIP = GetLocalIP
        
        strOutPutExeStep = "取客户机IP"
    End If

    For i = 0 To UBound(arrXML)
        'XML开始
        strXML_In = "<ROOT>"
        
        '业务操作信息
        strXML_In = strXML_In & vbCrLf & GetXMLFormat("OPSYSTEM", "HIS")
        strXML_In = strXML_In & vbCrLf & GetXMLFormat("OPWINID", IIf(intOprId = 0, "", intOprId))
        strXML_In = strXML_In & vbCrLf & GetXMLFormat("OPTYPE", intType)
        strXML_In = strXML_In & vbCrLf & GetXMLFormat("OPIP", gstrIP)
        strXML_In = strXML_In & vbCrLf & GetXMLFormat("OPMANNO", strUserCode)
        strXML_In = strXML_In & vbCrLf & GetXMLFormat("OPMANNAME", strUserName)
    
        '业务操作主信息
        strXML_In = strXML_In & vbCrLf & CStr(arrXML(i))
        
        'XML结束标志
        strXML_In = strXML_In & vbCrLf & "</ROOT>"
        
        strOutPutExeStep = "上传数据" & i + 1 & "/" & UBound(arrXML) + 1
        
        '输出上传数据
        If GBLN_OUTPUTLOG_DETAIL = True Then strOutput = strOutput & vbCrLf & strXML_In
       
        '调用接口方法上传数据
        strXML_Out = gobjSOAP.HisTransData(strXML_In)
        
        strOutPutExeStep = "调用对方接口完成"
        
        strOutput = strOutput & vbCrLf & "返回信息" & vbCrLf & strXML_Out
        
        '去掉回车换行符
        strXML_Out = Replace(strXML_Out, vbCrLf, "")
        strXML_Out = Replace(strXML_Out, vbCr, "")
        strXML_Out = Replace(strXML_Out, vbLf, "")
        
        '解析返回参数
        strTmp = strXML_Out
        strTmp = Mid(strTmp, 1, InStr(1, strTmp, "</RETVAL>") - 1)
        strOut_RETVAL = Mid(strTmp, InStr(1, strTmp, "<RETVAL>") + Len("<RETVAL>"))
        
        strTmp = strXML_Out
        strTmp = Mid(strTmp, 1, InStr(1, strTmp, "</RETMSG>") - 1)
        strOut_RETMSG = Mid(strTmp, InStr(1, strTmp, "<RETMSG>") + Len("<RETMSG>"))
        
        strTmp = strXML_Out
        strTmp = Mid(strTmp, 1, InStr(1, strTmp, "</RETCODE>") - 1)
        strOut_RETCODE = Mid(strTmp, InStr(1, strTmp, "<RETCODE>") + Len("<RETCODE>"))
               
        strOutPutExeStep = "解析返回参数完成"
               
        '返回1表示接口调用成功，其他值为不成功
        If strOut_RETCODE <> "1" Then
            If gblnShowMsg Then
                MsgBox strOut_RETMSG, vbInformation + vbOKOnly, GSTR_MESSAGE
            Else
                strReturn = strOut_RETMSG
            End If
            If blnShow Then frmDYEY_MZ_TransDrug.UnloadMe
            
            strOutput = strOutput & vbCrLf & "上传数据错误！"
            If GBLN_OUTPUTLOG_DETAIL = False Then strOutput = strOutput & vbCrLf & "最后一次上传数据" & vbCrLf & CStr(arrXML(i))
            strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransData_CQFLQZYY"
            Call OutputLog(strOutput)
    
            Exit Function
        End If
        
        If intType = gType.IntDrug Or intType = gType.IntStore Or intType = gType.IntDept Then
            If i = 0 Then
                frmDYEY_MZ_TransDrug.Show
                blnShow = True
                
                strOutPutExeStep = "打开上传药品信息窗口"
            End If
            
            Call frmDYEY_MZ_TransDrug.ChangePrg(i + 1, UBound(arrXML) + 1, intType)
            
            strOutPutExeStep = "上传药品信息窗口进度条执行"
        ElseIf intType = gType.IntDetail Then
            '根据上传信息取库房ID
            lngDrugStockID = GetStockID(arrXML(i))
            
            strOutPutExeStep = "取库房ID"
            
            If Not SetSendWin(lngDrugStockID, strNO, Val(strOut_RETMSG)) Then
                If gblnShowMsg Then
                    MsgBox "调整处方的发药窗口失败！", vbCritical, GSTR_MESSAGE
                Else
                    strReturn = "调整处方的发药窗口失败！"
                End If
                
                strOutPutExeStep = "调整处方的发药窗口失败！（" & lngDrugStockID & "；" & strNO & "；" & strOut_RETMSG & "）"
                strOutput = strOutput & vbCrLf & "调整窗口失败：（库房：" & lngDrugStockID & "；NO：" & strNO & "；返回消息：" & strOut_RETMSG & "）"
            Else
                strOutPutExeStep = "调整处方的发药窗口成功！（" & lngDrugStockID & "；" & strNO & "；" & strOut_RETMSG & "）"
                strOutput = strOutput & vbCrLf & "调整窗口成功：（库房：" & lngDrugStockID & "；NO：" & strNO & "；返回消息：" & strOut_RETMSG & "）"
            End If
            
        End If
    Next
    
    If intType = gType.IntDrug Or intType = gType.IntStore Or intType = gType.IntDept Then
        If gblnShowMsg Then
            MsgBox "上传完成！", vbInformation, GSTR_MESSAGE
        Else
            strReturn = "上传完成！"
        End If
    End If
    
    DYEY_MZ_TransData_CQFLQZYY = True
        
    strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransData_CQFLQZYY"
   
    Exit Function
errHandle:
    If gblnShowMsg Then
        If gintMode = 0 Then
            If gobjComLib.ErrCenter = 1 Then Resume
            Call gobjComLib.SaveErrLog
        Else
            MsgBox Err.Description, vbInformation, GSTR_SYSNAME
        End If
    End If
    
    strOutput = strOutput & vbCrLf & "发生异常错误"
    strOutput = strOutput & vbCrLf & "最后步骤：" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransData_CQFLQZYY"
    Call OutputLog(strOutput)
End Function


Public Function GetXML_Drug() As Variant
'将药品基础信息组织成指定的XML格式
'适用接口：韦乐海茨CONSIS系统v2.2
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim strTitle As String
    Dim arrXML As Variant
    Dim strErrMsg As String
    
    On Error GoTo errHandle
'    MsgBox "获取数据"
    strErrMsg = "获取数据"
    mstrSQL = "Select Distinct a.id 药品编号, a.名称 药品名称, e.名称 药品商品名, a.规格 药品规格, a.规格 药品包装规格, b.门诊单位 药品单位," & vbNewLine & _
              "    round(b.药库包装/b.门诊包装, 2) 包装比,b.门诊可否分零,a.产地 药品厂家, c.现价 * b.门诊包装 药品价格, d.药品剂型, " & vbNewLine & _
              "    b.门诊包装, a.建档时间 最后更新时间, f.简码 药品拼音, d.毒理分类 " & vbNewLine & _
              "From 收费项目目录 a, 药品规格 b, 收费价目 c, 药品特性 d, 收费项目别名 e, 收费项目别名 f " & vbNewLine & _
              "Where a.Id = b.药品id And a.Id = c.收费细目id And b.药名id = d.药名id And a.Id = e.收费细目id(+) And a.Id = f.收费细目id(+) And " & vbNewLine & _
              "    (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And Sysdate Between c.执行日期 And " & vbNewLine & _
              "    Nvl(c.终止日期, Sysdate) And e.性质(+) = 3 And f.性质(+) = 1 And f.码类(+) = 1"
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_Drug")
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_Drug")
    End If
    strErrMsg = "数据获取完毕"
    strXML = ""
    arrXML = Array()
    
    strErrMsg = "XML开始"
    With rsTemp
        If .RecordCount > 0 Then
            strTitle = "<ROOT>"
            
            Do While Not .EOF
                strDrug = "<CONSIS_BASIC_DRUGSVW"
                strDrug = strDrug & vbCrLf & "DRUG_CODE = """ & SpecialChar(!药品编号) & """"
                strDrug = strDrug & vbCrLf & "DRUG_NAME = """ & SpecialChar(!药品名称) & """"
                strDrug = strDrug & vbCrLf & "TRADE_NAME = """ & SpecialChar(!药品商品名) & """"
                strDrug = strDrug & vbCrLf & "DRUG_SPEC = """ & SpecialChar(!药品规格) & """"
                strDrug = strDrug & vbCrLf & "DRUG_PACKAGE = """ & NVL(!门诊包装) & """"  ' & SpecialChar(!药品包装规格) & """"
                strDrug = strDrug & vbCrLf & "DRUG_UNIT = """ & SpecialChar(!药品单位) & """"
                strDrug = strDrug & vbCrLf & "FIRM_ID = """ & SpecialChar(!药品厂家) & """"
                strDrug = strDrug & vbCrLf & "DRUG_PRICE = """ & NVL(!药品价格) & """"
                strDrug = strDrug & vbCrLf & "DRUG_FORM = """ & SpecialChar(!药品剂型) & """"
                strDrug = strDrug & vbCrLf & "DRUG_SORT = """ & SpecialChar(!毒理分类) & """"
                strDrug = strDrug & vbCrLf & "BARCODE = """""
                strDrug = strDrug & vbCrLf & "LAST_DATE = """ & Format(!最后更新时间, "yyyy-MM-DDThh:mm:ss") & """"
                strDrug = strDrug & vbCrLf & "PINYIN = """ & SpecialChar(!药品拼音) & """"
                strDrug = strDrug & vbCrLf & "DRUG_CONVERTATION = """ & NVL(!包装比) & """"
                strDrug = strDrug & vbCrLf & ">"
                strDrug = strDrug & vbCrLf & "</CONSIS_BASIC_DRUGSVW>"
                
                If Len(strXML & strDrug) > 3900 Then
                    '将以前的添加到数组
                    strXML = strXML & vbCrLf & "</ROOT>"
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                    strErrMsg = "装入数据1"
                    '重新拼凑新的XML
                    strXML = strTitle & vbCrLf & strDrug
                Else
                    strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
                End If
                
                rsTemp.MoveNext
                If .EOF And strXML <> "" Then
                    strXML = strXML & vbCrLf & "</ROOT>"
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                    strErrMsg = "装入数据2"
                End If
            Loop
        End If
    End With
    
    strErrMsg = "获取数据"
    GetXML_Drug = arrXML
    strErrMsg = "返回数据"
    Exit Function

errHandle:
    Debug.Print strErrMsg
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
End Function

Public Function GetXML_Drug_CQFLQZYY(ByRef strOutput As String) As Variant
'将药品基础信息组织成指定的XML格式
'适用接口：韦乐海茨CONSIS系统v4.3
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim arrXML As Variant
    Dim lngCount As Long
    Dim strOutPutExeStep As String    '执行步骤，用于输出日志方便查找问题
    
    '接口数据格式
    '字段名      类型            说明       NULL
    'Drug_code   Nvarchar(200)   药品编号    N
    'Drug_name   Nvarchar(200)   药品别名    N
    'Tradename   Nvarchar(200)   商品名称    Y
    'Englishname Nvarchar(200)   药品英文名  Y
    'Pinyin  Nvarchar(1000)  商品拼音码  Y
    'SortType1   Nvarchar(40)    药品类别    Y
    'SortType2   Nvarchar(40)    药品剂型    Y
    'Drug_spec   Nvarchar(200)   药品规格    N
    'MinSpecs    Nvarchar(200)   药品最小规格    Y
    'Unit    Nvarchar(40)    包装单位    N
    'MaxUNIT Nvarchar(40)    大包装单位  N
    'MinUNIT Nvarchar(40)    最小单位    N
    'Dosage  Numeric(20,6)   最小单位剂量    N
    'DosageUnit  Nvarchar(40)    剂量单位    Y
    'Price1  Numeric(20,6)   药品价格    N
    'Convertion1 Numeric(10,0)   大包装单位到包装单位换算率  N
    'Convertion2 Numeric(10,0)   包装单位到小包装单位换算率  N
    'Firm_id Nvarchar(200)   生产厂家编码    Y
    'Firm_name   Nvarchar(200)   生产厂家名称    Y
    'Passno  Nvarchar(200)   批准文号/注册证号   Y
    'BarCode Nvarchar(200)   药品条码    Y
    'StorageCondition    Nvarchar(200)   储存条件    Y
    'Storagetype Char(1) 储存类型(默认'0')   N
    'Allowind    Char(1) 停用标志（Y/N）
    '                       'Y'启用
    '                       'N'停用 N
    
    strOutput = strOutput & vbCrLf & "调用函数：GetXML_Drug_CQFLQZYY"
    
    On Error GoTo errHandle
              
    mstrSQL = "Select Distinct a.编码 As 药品编号, a.名称 As 药品名称, e.名称 As 药品商品名, g.名称 As 英文名, f.简码 As 拼音码," & vbNewLine & _
        " Decode(a.类别, 5, '西药', 6, '成药', '草药') As 类别, d.药品剂型, a.规格 As 药品规格, b.门诊单位 As 包装单位, b.药库单位 As 大包装单位," & vbNewLine & _
        " a.计算单位 As 最小单位, b.剂量系数, i.计算单位 As 剂量单位, c.现价 * b.门诊包装 As 药品价格, Round(b.药库包装 / b.门诊包装, 2) As 大包装系数," & vbNewLine & _
        " b.门诊包装 As 包装系数, j.编码 As 生产厂家编码, j.名称 As 生产厂家名称, b.上次批准文号 As 批准文号," & vbNewLine & _
        " Decode(Nvl(a.撤档时间, To_Date('3000-01-01', 'yyyy-MM-dd')), To_Date('3000-01-01', 'yyyy-MM-dd'), 'Y', 'N') As 停用标志" & vbNewLine & _
        " From 收费项目目录 A, 药品规格 B, 收费价目 C, 药品特性 D, 收费项目别名 E, 收费项目别名 F, 收费项目别名 G, 诊疗项目目录 I, 药品生产商 J" & vbNewLine & _
        " Where a.Id = b.药品id And a.Id = c.收费细目id And b.药名id = d.药名id And a.Id = e.收费细目id(+) And a.Id = f.收费细目id(+) And" & vbNewLine & _
        " Sysdate Between c.执行日期 And Nvl(c.终止日期, Sysdate) And e.性质(+) = 3 And f.性质(+) = 1 And f.码类(+) = 1 And" & vbNewLine & _
        " a.Id = g.收费细目id(+) And g.性质(+) = 2 And b.药名id = i.Id And b.上次产地 = j.名称(+)"
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_Drug_CQFLZYY")
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_Drug_CQFLZYY")
    End If
    
    strOutPutExeStep = "执行SQL成功"
    
    strXML = ""
    arrXML = Array()

    With rsTemp
        strOutPutExeStep = "拼装XML：begin"
        
        If .RecordCount > 0 Then
            Do While Not .EOF
                strDrug = "<CONSIS_BASIC_DRUGSVW>"
               
                strDrug = strDrug & vbCrLf & GetXMLFormat("DRUG_CODE", SpecialChar(!药品编号))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DRUG_NAME", SpecialChar(!药品名称))
                strDrug = strDrug & vbCrLf & GetXMLFormat("TRADENAME", SpecialChar(!药品商品名))
                strDrug = strDrug & vbCrLf & GetXMLFormat("ENGLISHNAME", SpecialChar(!英文名))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PINYIN", SpecialChar(!拼音码))
                
                strOutPutExeStep = "拼装XML：1"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("SORTTYPE1", SpecialChar(!类别))
                strDrug = strDrug & vbCrLf & GetXMLFormat("SORTTYPE2", SpecialChar(!药品剂型))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DRUG_SPEC", SpecialChar(!药品规格))
                strDrug = strDrug & vbCrLf & GetXMLFormat("MINSPECS", "")
                strDrug = strDrug & vbCrLf & GetXMLFormat("UNIT", SpecialChar(NVL(!包装单位)))
                
                strOutPutExeStep = "拼装XML：2"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("MAXUNIT", SpecialChar(NVL(!大包装单位)))
                strDrug = strDrug & vbCrLf & GetXMLFormat("MINUNIT", SpecialChar(NVL(!最小单位)))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DOSAGE", NVL(!剂量系数))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DOSAGEUNIT", SpecialChar(NVL(!剂量单位)))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRICE1", NVL(!药品价格))
                
                strOutPutExeStep = "拼装XML：3"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("CONVERTION1", NVL(!大包装系数))
                strDrug = strDrug & vbCrLf & GetXMLFormat("CONVERTION2", NVL(!包装系数))
                strDrug = strDrug & vbCrLf & GetXMLFormat("FIRM_ID", SpecialChar(NVL(!生产厂家编码)))
                strDrug = strDrug & vbCrLf & GetXMLFormat("FIRM_NAME", SpecialChar(NVL(!生产厂家名称)))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PASSNO", SpecialChar(NVL(!批准文号)))
                
                strOutPutExeStep = "拼装XML：4"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("BARCODE", "")
                strDrug = strDrug & vbCrLf & GetXMLFormat("STORAGECONDITION", "")
                strDrug = strDrug & vbCrLf & GetXMLFormat("STORAGETYPE", "")
                strDrug = strDrug & vbCrLf & GetXMLFormat("ALLOWIND", NVL(!停用标志))
                
                strOutPutExeStep = "拼装XML：5"
                
                strDrug = strDrug & vbCrLf & "</CONSIS_BASIC_DRUGSVW>"
                
                lngCount = lngCount + 1
                
                '每500个药品组成一个数组分批上传
                If lngCount > 500 Then
                    '将以前的添加到数组
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                    '重新拼凑新的XML
                    strXML = strDrug
                    lngCount = 0
                    
                    strOutPutExeStep = "拼装XML：7"
                Else
                    strXML = IIf(strXML = "", "", strXML & vbCrLf) & strDrug
                    
                    strOutPutExeStep = "拼装XML：6"
                End If
                
                rsTemp.MoveNext
                
                If .EOF And strXML <> "" Then
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                    
                    strOutPutExeStep = "拼装XML：end"
                End If
            Loop
        End If
    End With
    
    GetXML_Drug_CQFLQZYY = arrXML
    
    strOutput = strOutput & vbCrLf & "执行成功：GetXML_Drug_CQFLQZYY"
  
    Exit Function
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
    
    strOutput = strOutput & vbCrLf & "发生异常错误"
    strOutput = strOutput & vbCrLf & "最后SQL" & vbCrLf & mstrSQL
    strOutput = strOutput & vbCrLf & "最后步骤：" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "执行失败：GetXML_Drug_CQFLQZYY"
    Call OutputLog(strOutput)
End Function

Public Function GetXML_RecipeDetail(ByVal strStockIDs As String, ByVal strNO As String) As Variant
'将处方明细组织成指定的XML格式
'适用接口：韦乐海茨CONSIS系统v2.2
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim strTitle As String
    Dim arrXML As Variant
    Dim strSQL As String
    Dim i As Integer
    Dim rsDetails As Recordset
    Dim strDetail As String
    
    Call OutputLog("进入GetXML_RecipeDetail")
    
    On Error GoTo errHandle
    '获取处方单信息
    strSQL = "Select a.填制日期 处方时间, a.单据, a.No 处方编号, a.库房id 发药药局, c.病人id 就诊卡号, a.姓名 患者姓名, Decode(a.优先级, 1, '01', '00') 患者类型, " & vbNewLine & _
             "    c.出生日期 患者出生日期, c.性别 患者性别, c.身份 患者身份, c.医疗付款方式 医保类型, Sum(d.应收金额) 费用, Sum(d.实收金额) 实付费用," & vbNewLine & _
             "    f.id 开单科室, d.开单人 开方医生, d.开单人 录方人, Decode(a.优先级, 1, '1', '2') 配药优先级 " & vbNewLine & _
             "From 未发药品记录 a, 病人信息 c, 门诊费用记录 d, 药品收发记录 e, 部门表 f " & vbNewLine & _
             "  ,Table(Cast(f_Str2list2([2], '|', ',') As t_Strlist2)) I " & vbNewLine & _
             "Where a.单据 = e.单据 And a.No = e.No And a.库房id = e.库房id And a.病人id = c.病人id And e.费用id = d.Id And " & vbNewLine & _
             "    d.开单部门id = f.Id And a.单据 = i.c1 And a.NO = i.c2 " & _
             IIf(Trim(strStockIDs) = "", "", " And Instr([1], ';' || a.库房id || ';') > 0 ")

    strSQL = strSQL & _
             "Group By a.填制日期, a.单据, a.No, a.库房id, c.病人id, a.姓名, Decode(a.优先级, 1, '01', '00'), c.出生日期, c.性别, " & vbNewLine & _
             "    c.身份, c.医疗付款方式, f.id, d.开单人,d.开单人, Decode(a.优先级, 1, '1', '2') "

    mstrSQL = strSQL & vbCrLf & " union all  " & vbCrLf & Replace(strSQL, "门诊费用记录", "住院费用记录")
    mstrSQL = "select * from (" & mstrSQL & ") Order By 发药药局, 就诊卡号 "
    
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_RecipeDetail", ";" & strStockIDs & ";", strNO)
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_RecipeDetail", ";" & strStockIDs & ";", strNO)
    End If
    
    '获取处方明细信息
    strSQL = "Select Distinct a.填制日期, a.单据, a.No, a.序号, b.id 药品编码, b.名称 药品名称, c.名称 药品商品名, b.规格 药品规格, b.规格 药品包装规格, " & vbNewLine & _
             "    d.门诊单位 药品单位, a.产地 药品厂家, a.零售价 * d.门诊包装 药品价格, a.实际数量 / d.门诊包装 数量, e.应收金额 费用,e.病人id," & vbNewLine & _
             "    e.实收金额 实付费用, a.单量 药品剂量, a.库房id, a.用法, f.执行频次, g.计算单位 剂量单位 " & vbNewLine & _
             "From 药品收发记录 a, 收费项目目录 b, 收费项目别名 c, 药品规格 d, 门诊费用记录 e, 病人医嘱记录 f, 诊疗项目目录 g " & vbNewLine & _
             "  ,Table(Cast(f_Str2list2([2], '|', ',') As t_Strlist2)) I " & vbNewLine & _
             "Where a.药品id = b.Id And a.药品id = c.收费细目id(+) And a.药品id = d.药品id And a.费用id = e.Id and d.药名id=g.id " & vbNewLine & _
             "    And e.医嘱序号 = f.Id(+) And c.性质(+) = 3 And a.单据 = i.c1 And a.NO = i.c2 " & _
             IIf(Trim(strStockIDs) = "", "", " And Instr([1], ';' || a.库房id || ';') > 0 ")

    mstrSQL = strSQL & vbCrLf & " union all  " & vbCrLf & Replace(strSQL, "门诊费用记录", "住院费用记录")
Call OutputLog("查询处方明细开始")
    If gintMode = 0 Then
        Set rsDetails = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_RecipeDetail", ";" & strStockIDs & ";", strNO)
    Else
        Set rsDetails = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_RecipeDetail", ";" & strStockIDs & ";", strNO)
    End If
    strXML = ""
    arrXML = Array()
    
Call OutputLog("查询处方明细完成。")
    
'    '库房ID为0的情况单独函数处理
'    If lngStockID = 0 Then
'Call OutputLog("执行GetXML_RecipeDetailEx")
'        If GetXML_RecipeDetailEx(rsTemp, rsDetails, arrXML) Then
'            GetXML_RecipeDetail = arrXML
'        End If
'        Exit Function
'    End If
    
    With rsTemp
        If .RecordCount > 0 Then
            strTitle = "<ROOT>"
            
            Do While Not .EOF
                strDrug = "<CONSIS_PRESC_MSTVW"
                strDrug = strDrug & vbCrLf & "PRESC_DATE = """ & Format(!处方时间, "yyyy-MM-DDThh:mm:ss") & """"
                strDrug = strDrug & vbCrLf & "PRESC_NO = """ & SpecialChar(!处方编号) & """"
                strDrug = strDrug & vbCrLf & "DISPENSARY = """ & NVL(!发药药局) & """"
                strDrug = strDrug & vbCrLf & "PATIENT_ID = """ & NVL(!就诊卡号) & """"
                strDrug = strDrug & vbCrLf & "PATIENT_NAME = """ & SpecialChar(!患者姓名) & """"
                strDrug = strDrug & vbCrLf & "PATIENT_TYPE = """ & NVL(!患者类型) & """"
                strDrug = strDrug & vbCrLf & "DATE_OF_BIRTH = """ & Format(NVL(!患者出生日期), "yyyy-MM-DDThh:mm:ss") & """"
                strDrug = strDrug & vbCrLf & "SEX = """ & SpecialChar(!患者性别) & """"
                strDrug = strDrug & vbCrLf & "PRESC_IDENTITY = """ & SpecialChar(!患者身份) & """"
                strDrug = strDrug & vbCrLf & "CHARGE_TYPE = """ & SpecialChar(!医保类型) & """"
                strDrug = strDrug & vbCrLf & "PRESC_ATTR = """""
                strDrug = strDrug & vbCrLf & "PRESC_INFO = """""
                strDrug = strDrug & vbCrLf & "RCPT_INFO = " & GetRCPT_INFO(NVL(!处方编号))
                strDrug = strDrug & vbCrLf & "RCPT_REMARK = """""
                strDrug = strDrug & vbCrLf & "REPETITION = ""1"""
                strDrug = strDrug & vbCrLf & "COSTS = """ & NVL(!费用) & """"
                strDrug = strDrug & vbCrLf & "PAYMENTS = """ & NVL(!实付费用) & """"
                strDrug = strDrug & vbCrLf & "ORDERED_BY = """ & NVL(!开单科室) & """"
                strDrug = strDrug & vbCrLf & "PRESCRIBED_BY = """ & SpecialChar(!开方医生) & """"
                strDrug = strDrug & vbCrLf & "ENTERED_BY = """ & SpecialChar(!录方人) & """"
                strDrug = strDrug & vbCrLf & "DISPENSE_PRI = """ & NVL(!配药优先级) & """"
                strDrug = strDrug & vbCrLf & ">"
                
                '过滤明细记录，确保与单据对应
                rsDetails.Filter = "no='" & !处方编号 & "' and 单据=" & NVL(!单据) & " and 库房id=" & NVL(!发药药局)
                rsDetails.Sort = "序号"
                
                strDetail = ""
                Do While Not rsDetails.EOF
                    strDetail = strDetail & vbCrLf & "<CONSIS_PRESC_DTLVW"
                    strDetail = strDetail & vbCrLf & "PRESC_DATE = """ & Format(rsDetails!填制日期, "yyyy-MM-DDThh:mm:ss") & """"
                    strDetail = strDetail & vbCrLf & "PRESC_NO = """ & NVL(rsDetails!no) & """"
                    strDetail = strDetail & vbCrLf & "ITEM_NO = """ & NVL(rsDetails!序号) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_CODE = """ & SpecialChar(rsDetails!药品编码) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_NAME = """ & SpecialChar(rsDetails!药品名称) & """"
                    strDetail = strDetail & vbCrLf & "TRADE_NAME = """ & SpecialChar(rsDetails!药品商品名) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_SPEC= """ & SpecialChar(rsDetails!药品规格) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_PACKAGE = """ & SpecialChar(rsDetails!药品包装规格) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_UNIT = """ & SpecialChar(rsDetails!药品单位) & """"
                    strDetail = strDetail & vbCrLf & "FIRM_ID = """ & SpecialChar(rsDetails!药品厂家) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_PRICE = """ & NVL(rsDetails!药品价格) & """"
                    strDetail = strDetail & vbCrLf & "QUANTITY = """ & NVL(rsDetails!数量) & """"
                    strDetail = strDetail & vbCrLf & "COSTS = """ & NVL(rsDetails!费用) & """"
                    strDetail = strDetail & vbCrLf & "PAYMENTS = """ & NVL(rsDetails!实付费用) & """"
                    strDetail = strDetail & vbCrLf & "DOSAGE = """ & NVL(rsDetails!药品剂量) & """"
                    strDetail = strDetail & vbCrLf & "DOSAGE_UNITS = """ & SpecialChar(rsDetails!剂量单位) & """"
                    strDetail = strDetail & vbCrLf & "ADMINISTRATION = """ & SpecialChar(rsDetails!用法) & """"
                    strDetail = strDetail & vbCrLf & "FREQUENCY = """ & SpecialChar(rsDetails!执行频次) & """"
                    strDetail = strDetail & vbCrLf & ">"
                    strDetail = strDetail & vbCrLf & "</CONSIS_PRESC_DTLVW>"
                    rsDetails.MoveNext
                Loop
                strDrug = strDrug & strDetail
                strDrug = strDrug & vbCrLf & "</CONSIS_PRESC_MSTVW>"
                
                strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
                rsTemp.MoveNext
                If .EOF Then
                    strXML = strXML & vbCrLf & "</ROOT>"
                    
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
Call OutputLog(strXML)
                End If
            Loop
        End If
    End With
    
    GetXML_RecipeDetail = arrXML
    Exit Function
    
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        'MsgBox Err.Description, vbInformation, GSTR_SYSNAME
        Call OutputLog("出错变量值：" & strDrug & vbCr & strDetail)
    End If
End Function

Public Function GetXML_RecipeDetail_CQFLQZYY(ByVal strStockIDs As String, ByVal strNO As String, ByRef strOutput As String) As Variant
'将处方明细组织成指定的XML格式
'适用接口：韦乐海茨CONSIS系统v4.3
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim arrXML As Variant
    Dim strSQL As String
    Dim i As Integer
    Dim rsDetails As Recordset
    Dim strDetail As String
    Dim strOutPutExeStep As String    '执行步骤，用于输出日志方便查找问题
    
    '处方主表
'    英文标识    中文标识    数据类型    Nullable
'    Presc_date  处方时间    Datetime    N
'    Presc_no    处方编号    Nvarchar(200)   N
'    Dispensary  发药药局编号    Nvarchar(40)    N
'    Patient_id  就诊卡号    Nvarchar(40)    N
'    Patient_name    患者姓名    Nvarchar(200)   N
'    Invoice_no  发票编号    Nvarchar(200)   Y
'    Patient_type 患者类型
'    '00' 普通
'    '01' 特需   Nvarchar(40)    Y
'    Date_of_birth   患者出生日期    Datetime    N
'    Sex 患者性别(男/女) Nvarchar(40)    N
'    Presc_identity  患者身份    Nvarchar(40)    Y
'    Charge_type 医保类型    Nvarchar(40)    Y
'    Presc_attr 处方属性
'    手工处方，临时处方等文本信息    Nvarchar(1000)  Y
'    Presc_info 处方类型
'    费用相关处方类型文本信息（计费方式）    Nvarchar(1000)  Y
'    Rcpt_info   诊断信息    Nvarchar(1000)  Y
'    Rcpt_remark 处方备注信息    Nvarchar(1000)  Y
'    Repetition  剂数    Numeric(10,0)   N
'    Costs   费用    Numeric(20,6)   N
'    Payments    实付费用    Numeric(20,6)   N
'    Ordered_by  开单科室编号    Nvarchar(40)    Y
'    Ordered_by_name 开单科室名称    Nvarchar(40)    Y
'    Prescribed_by   开方医生    Nvarchar(40)    Y
'    Entered_by  录方人  Nvarchar(40)    Y
'    Dispense_pri    药优先级（付费处到药房距离）数字从小到大表示    Numeric(10,0)   Y
    
    '处方明细
'    英文标识    中文标识    数据类型    Nullable
'    Presc_no    处方编号    Nvarchar(200)   N
'    Item_no 药品序号    Numeric(10,0)   N
'    Advice_code 医嘱编号    Nvarchar(200)   Y
'    Drug_code   药品编号    Nvarchar(200)   N
'    Drug_spec   药品规格    Nvarchar(200)   Y
'    Drug_name   药品名称    Nvarchar(200)   N
'    Firm_id 厂商编号    Nvarchar(200)   Y
'    Firm_name   厂商名称    Nvarchar(200)   Y
'    Package_spec    药品包装规格    Nvarchar(200)   Y
'    Package_units   药品包装单位    Nvarchar(40)    Y
'    Quantity    数量    Numeric(20,6)   N
'    Unit    药品单位    Nvarchar(40)    N
'    Costs   费用    Numeric(20,6)   N
'    Payments    实付费用    Numeric(20,6)   N
'    Dosage  药品剂量（每次服用量）  Nvarchar(40)    Y
'    Dosage_units    剂量单位（每次服用单位）    Nvarchar(40)    Y
'    Administration  药品用法（使用方法）    Nvarchar(200)   Y
'    frequency   药品用量（使用频率 每天几次）   Nvarchar(200)   Y
'    Additionusage   补充用法    Nvarchar(200)   Y
'    Rcpt_remark 处方明细备注信息    Nvarchar(1000)  Y
    
    On Error GoTo errHandle
    
    strOutput = strOutput & vbCrLf & "调用函数：GetXML_RecipeDetail_CQFLQZYY"
    
    '获取处方单信息
    strSQL = "Select a.填制日期 处方时间, a.单据, a.No 处方编号, a.库房id 发药药局, c.病人id 就诊卡号, a.姓名 患者姓名, Decode(a.优先级, 1, '01', '00') 患者类型, " & vbNewLine & _
             "    c.出生日期 患者出生日期, c.性别 患者性别, c.身份 患者身份, c.医疗付款方式 医保类型, Sum(d.应收金额) 费用, Sum(d.实收金额) 实付费用," & vbNewLine & _
             "    f.编码 开单科室编码,f.名称 As 开单科室名称, d.开单人 开方医生, d.开单人 录方人, Decode(a.优先级, 1, '1', '2') 配药优先级 " & vbNewLine & _
             "From 未发药品记录 a, 病人信息 c, 门诊费用记录 d, 药品收发记录 e, 部门表 f " & vbNewLine & _
             "  ,Table(Cast(f_Str2list2([2], '|', ',') As t_Strlist2)) I "
    
    '包头市中心医院要传住院的麻醉和精神药品数据
    If gstrUnit = GCST_UNIT_BTSZXYY Then
        strSQL = strSQL & " , 药品规格 G, 药品特性 T "
    End If
             
    strSQL = strSQL & "Where a.单据 = e.单据 And a.No = e.No And a.库房id = e.库房id And a.病人id = c.病人id And e.费用id = d.Id And " & vbNewLine & _
             "    d.开单部门id = f.Id And a.单据 = i.c1 And a.NO = i.c2 " & _
             IIf(Trim(strStockIDs) = "", "", " And Instr([1], ';' || a.库房id || ';') > 0 ")
    
    '包头市中心医院要传住院的麻醉和精神药品数据
    If gstrUnit = GCST_UNIT_BTSZXYY Then
        strSQL = strSQL & " And e.药品id = g.药品id And g.药名id = t.药名id And (a.单据 = 9 And t.毒理分类 In ('麻醉药', '精神I类') Or a.单据 = 8) "
    End If
    
    strSQL = strSQL & _
             "Group By a.填制日期, a.单据, a.No, a.库房id, c.病人id, a.姓名, Decode(a.优先级, 1, '01', '00'), c.出生日期, c.性别, " & vbNewLine & _
             "    c.身份, c.医疗付款方式, f.编码, f.名称, d.开单人,d.开单人, Decode(a.优先级, 1, '1', '2') "
             

    mstrSQL = strSQL & vbCrLf & " union all  " & vbCrLf & Replace(strSQL, "门诊费用记录", "住院费用记录")
    mstrSQL = "select * from (" & mstrSQL & ") Order By 发药药局, 就诊卡号 "
    
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_RecipeDetail_CQFLQZYY", ";" & strStockIDs & ";", strNO)
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_RecipeDetail_CQFLQZYY", ";" & strStockIDs & ";", strNO)
    End If
    
    strOutPutExeStep = "执行处方单信息SQL成功"
    
    '获取处方明细信息
    strSQL = "Select Distinct a.填制日期, a.单据, a.No, a.序号, b.编码 药品编码, b.名称 药品名称, c.名称 药品商品名, b.规格 药品规格, b.规格 药品包装规格, " & vbNewLine & _
             "    d.门诊单位 药品单位, h.编码 as 生产厂家编码,a.产地 生产厂家名称, a.零售价 * d.门诊包装 药品价格, a.实际数量 / d.门诊包装 数量, e.应收金额 费用,e.病人id," & vbNewLine & _
             "    e.实收金额 实付费用,e.医嘱序号 , a.单量 药品剂量, a.库房id, a.用法, f.执行频次, g.计算单位 剂量单位 " & vbNewLine & _
             "From 药品收发记录 a, 收费项目目录 b, 收费项目别名 c, 药品规格 d, 门诊费用记录 e, 病人医嘱记录 f, 诊疗项目目录 g, 药品生产商 h " & vbNewLine & _
             "  ,Table(Cast(f_Str2list2([2], '|', ',') As t_Strlist2)) I "
    
    '包头市中心医院要传住院的麻醉和精神药品数据
    If gstrUnit = GCST_UNIT_BTSZXYY Then
        strSQL = strSQL & " , 药品特性 T "
    End If
    
    strSQL = strSQL & " Where a.药品id = b.Id And a.药品id = c.收费细目id(+) And a.药品id = d.药品id And a.费用id = e.Id and d.药名id=g.id " & vbNewLine & _
             "    And e.医嘱序号 = f.Id(+) And c.性质(+) = 3 And a.产地=h.名称(+) And a.单据 = i.c1 And a.NO = i.c2 " & _
             IIf(Trim(strStockIDs) = "", "", " And Instr([1], ';' || a.库房id || ';') > 0 ")
    
    '包头市中心医院要传住院的麻醉和精神药品数据
    If gstrUnit = GCST_UNIT_BTSZXYY Then
        strSQL = strSQL & " And d.药名id = t.药名id And (a.单据 = 9 And t.毒理分类 In ('麻醉药', '精神I类') Or a.单据 = 8) "
    End If
    
    mstrSQL = strSQL & vbCrLf & " union all  " & vbCrLf & Replace(strSQL, "门诊费用记录", "住院费用记录")
    
    If gintMode = 0 Then
        Set rsDetails = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_RecipeDetail_CQFLQZYY", ";" & strStockIDs & ";", strNO)
    Else
        Set rsDetails = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_RecipeDetail_CQFLQZYY", ";" & strStockIDs & ";", strNO)
    End If
    
    strOutPutExeStep = "执行处方信息SQL成功"
    
    strXML = ""
    arrXML = Array()
    
    strOutPutExeStep = "拼装XML：begin"
    
    With rsTemp
        If .RecordCount > 0 Then
            Do While Not .EOF
                strDrug = "<CONSIS_PRESC_MSTVW>"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRESC_DATE", Format(!处方时间, "yyyy-MM-DDThh:mm:ss"))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRESC_NO", SpecialChar(!处方编号))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DISPENSARY", NVL(!发药药局))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PATIENT_ID", NVL(!就诊卡号))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PATIENT_NAME", SpecialChar(!患者姓名))
                
                strOutPutExeStep = "拼装处方单XML：1"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("INVOICE_NO", "")
                strDrug = strDrug & vbCrLf & GetXMLFormat("PATIENT_TYPE", NVL(!患者类型))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DATE_OF_BIRTH", Format(NVL(!患者出生日期), "yyyy-MM-DDThh:mm:ss"))
                strDrug = strDrug & vbCrLf & GetXMLFormat("SEX", SpecialChar(!患者性别))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRESC_IDENTITY", SpecialChar(!患者身份))
                
                strOutPutExeStep = "拼装处方单XML：2"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("CHARGE_TYPE", SpecialChar(!医保类型))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRESC_ATTR", "")
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRESC_INFO", "")
                strDrug = strDrug & vbCrLf & GetXMLFormat("RCPT_INFO", GetRCPT_INFO(NVL(!处方编号)))
                strDrug = strDrug & vbCrLf & GetXMLFormat("RCPT_REMARK", "")
                
                strOutPutExeStep = "拼装处方单XML：3"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("REPETITION", "1")
                strDrug = strDrug & vbCrLf & GetXMLFormat("COSTS", NVL(!费用))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PAYMENTS", NVL(!实付费用))
                strDrug = strDrug & vbCrLf & GetXMLFormat("ORDERED_BY", NVL(!开单科室编码))
                strDrug = strDrug & vbCrLf & GetXMLFormat("ORDERED_BY_NAME", NVL(!开单科室名称))
                
                strOutPutExeStep = "拼装处方单XML：4"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRESCRIBED_BY", SpecialChar(!开方医生))
                strDrug = strDrug & vbCrLf & GetXMLFormat("ENTERED_BY", SpecialChar(!录方人))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DISPENSE_PRI", NVL(!配药优先级))
                
                strOutPutExeStep = "拼装处方单XML：5"
                
                '过滤明细记录，确保与单据对应
                rsDetails.Filter = "no='" & !处方编号 & "' and 单据=" & NVL(!单据) & " and 库房id=" & NVL(!发药药局)
                rsDetails.Sort = "序号"
                
                strOutPutExeStep = "过滤明细记录"
                
                strDetail = ""
                Do While Not rsDetails.EOF
                    strDetail = strDetail & vbCrLf & "<CONSIS_PRESC_DTLVW>"
                    
                    strDetail = strDetail & vbCrLf & GetXMLFormat("PRESC_NO", NVL(rsDetails!no))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("ITEM_NO", NVL(rsDetails!序号))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("ADVICE_CODE", NVL(rsDetails!医嘱序号))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("DRUG_CODE", SpecialChar(rsDetails!药品编码))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("DRUG_SPEC", SpecialChar(rsDetails!药品规格))
                    
                    strOutPutExeStep = "拼装处方明细XML：1"
                    
                    strDetail = strDetail & vbCrLf & GetXMLFormat("DRUG_NAME", SpecialChar(rsDetails!药品名称))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("FIRM_ID", SpecialChar(rsDetails!生产厂家编码))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("FIRM_NAME", SpecialChar(rsDetails!生产厂家名称))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("PACKAGE_SPEC", SpecialChar(rsDetails!药品包装规格))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("PACKAGE_UNITS", SpecialChar(rsDetails!药品单位))
                    
                    strOutPutExeStep = "拼装处方明细XML：2"
                     
                    strDetail = strDetail & vbCrLf & GetXMLFormat("QUANTITY", NVL(rsDetails!数量))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("UNIT", SpecialChar(rsDetails!药品单位))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("COSTS", NVL(rsDetails!费用))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("PAYMENTS", NVL(rsDetails!实付费用))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("DOSAGE", NVL(rsDetails!药品剂量))
                    
                    strOutPutExeStep = "拼装处方明细XML：3"
                    
                    strDetail = strDetail & vbCrLf & GetXMLFormat("DOSAGE_UNITS", SpecialChar(rsDetails!剂量单位))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("ADMINISTRATION", SpecialChar(rsDetails!用法))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("FREQUENCY ", SpecialChar(rsDetails!执行频次))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("Additionusage", "")
                    strDetail = strDetail & vbCrLf & GetXMLFormat("Rcpt_remark", "")
                    
                    strOutPutExeStep = "拼装处方明细XML：4"
                     
                    strDetail = strDetail & vbCrLf & "</CONSIS_PRESC_DTLVW>"
                    rsDetails.MoveNext
                Loop
                strDrug = strDrug & strDetail
                strDrug = strDrug & vbCrLf & "</CONSIS_PRESC_MSTVW>"
                
                strXML = IIf(strXML = "", "", strXML & vbCrLf) & strDrug
                
                rsTemp.MoveNext
                
                If .EOF And strXML <> "" Then
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                    
                    strOutPutExeStep = "拼装XML：end"
                End If
            Loop
        End If
    End With
    
    GetXML_RecipeDetail_CQFLQZYY = arrXML
    
    strOutput = strOutput & vbCrLf & "执行成功：GetXML_RecipeDetail_CQFLQZYY"
    Call OutputLog(strOutput)
    
    Exit Function
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
    
    strOutput = strOutput & vbCrLf & "发生异常错误"
    strOutput = strOutput & vbCrLf & "最后SQL" & vbCrLf & mstrSQL
    strOutput = strOutput & vbCrLf & "最后步骤：" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "执行失败：GetXML_RecipeDetail_CQFLQZYY"
    Call OutputLog(strOutput)
End Function


Private Function GetXML_RecipeDetailEx(ByVal rsBill As ADODB.Recordset, ByVal rsDetail As ADODB.Recordset, ByRef varXML As Variant) As Boolean
'功能：处理库房ID为0的情况，分处理库房ID与病人ID生成XML字符串
'参数：
'  rsBill：单据数据集；
'  rsDetail：明细数据集；
'  varXML：生成的XML字符串数组（实参）。
'返回：True成功   False失败
'适用接口：韦乐海茨CONSIS系统v2.2
    Const STR_ROOT_BEGIN = "<ROOT>"
    Const STR_ROOT_END = "</ROOT>"
    Const STR_BILL = "CONSIS_PRESC_MSTVW"
    Const STR_DETAIL = "CONSIS_PRESC_DTLVW"
    Dim strXML As String, strBill As String, strDetail As String
    Dim lng库房ID As Long, lng病人ID As Long
    Dim varReturn As Variant
    
    On Error GoTo errHandle
    varReturn = Array()
    With rsBill
        If .RecordCount <= 0 Then Exit Function
        .MoveFirst
        lng库房ID = NVL(!发药药局, 0)
        lng病人ID = NVL(!就诊卡号, 0)
        Do
            If .EOF Then Exit Do
            '单据
            strBill = "<" & STR_BILL & " "
            strBill = strBill & vbCrLf & "PRESC_DATE = """ & Format(!处方时间, "yyyy-MM-DDThh:mm:ss") & """"
            strBill = strBill & vbCrLf & "PRESC_NO = """ & SpecialChar(!处方编号) & """"
            strBill = strBill & vbCrLf & "DISPENSARY = """ & NVL(!发药药局) & """"
            strBill = strBill & vbCrLf & "PATIENT_ID = """ & NVL(!就诊卡号) & """"
            strBill = strBill & vbCrLf & "PATIENT_NAME = """ & SpecialChar(!患者姓名) & """"
            strBill = strBill & vbCrLf & "PATIENT_TYPE = """ & NVL(!患者类型) & """"
            strBill = strBill & vbCrLf & "DATE_OF_BIRTH = """ & Format(NVL(!患者出生日期), "yyyy-MM-DDThh:mm:ss") & """"
            strBill = strBill & vbCrLf & "SEX = """ & SpecialChar(!患者性别) & """"
            strBill = strBill & vbCrLf & "PRESC_IDENTITY = """ & SpecialChar(!患者身份) & """"
            strBill = strBill & vbCrLf & "CHARGE_TYPE = """ & SpecialChar(!医保类型) & """"
            strBill = strBill & vbCrLf & "PRESC_ATTR = """""
            strBill = strBill & vbCrLf & "PRESC_INFO = """""
            strBill = strBill & vbCrLf & "RCPT_INFO = " & GetRCPT_INFO(NVL(!处方编号))
            strBill = strBill & vbCrLf & "RCPT_REMARK = """""
            strBill = strBill & vbCrLf & "REPETITION = ""1"""
            strBill = strBill & vbCrLf & "COSTS = """ & NVL(!费用) & """"
            strBill = strBill & vbCrLf & "PAYMENTS = """ & NVL(!实付费用) & """"
            strBill = strBill & vbCrLf & "ORDERED_BY = """ & NVL(!开单科室) & """"
            strBill = strBill & vbCrLf & "PRESCRIBED_BY = """ & SpecialChar(!开方医生) & """"
            strBill = strBill & vbCrLf & "ENTERED_BY = """ & SpecialChar(!录方人) & """"
            strBill = strBill & vbCrLf & "DISPENSE_PRI = """ & NVL(!配药优先级) & """"
            strBill = strBill & vbCrLf & ">"
            
            '过滤明细记录，确保与单据对应
            strDetail = ""
            rsDetail.Filter = "no='" & !处方编号 & "' and 单据=" & NVL(!单据) & " and 库房id=" & NVL(!发药药局) & " and 病人id=" & NVL(!就诊卡号)
            rsDetail.Sort = "序号"
            Do
                If rsDetail.EOF Then Exit Do
                '明细
                strDetail = strDetail & vbCrLf & "<" & STR_DETAIL & " "
                strDetail = strDetail & vbCrLf & "PRESC_DATE = """ & Format(rsDetail!填制日期, "yyyy-MM-DDThh:mm:ss") & """"
                strDetail = strDetail & vbCrLf & "PRESC_NO = """ & NVL(rsDetail!no) & """"
                strDetail = strDetail & vbCrLf & "ITEM_NO = """ & rsDetail!序号 & """"
                strDetail = strDetail & vbCrLf & "DRUG_CODE = """ & SpecialChar(rsDetail!药品编码) & """"
                strDetail = strDetail & vbCrLf & "DRUG_NAME = """ & SpecialChar(rsDetail!药品名称) & """"
                strDetail = strDetail & vbCrLf & "TRADE_NAME = """ & SpecialChar(rsDetail!药品商品名) & """"
                strDetail = strDetail & vbCrLf & "DRUG_SPEC= """ & SpecialChar(rsDetail!药品规格) & """"
                strDetail = strDetail & vbCrLf & "DRUG_PACKAGE = """ & SpecialChar(rsDetail!药品包装规格) & """"
                strDetail = strDetail & vbCrLf & "DRUG_UNIT = """ & SpecialChar(rsDetail!药品单位) & """"
                strDetail = strDetail & vbCrLf & "FIRM_ID = """ & SpecialChar(rsDetail!药品厂家) & """"
                strDetail = strDetail & vbCrLf & "DRUG_PRICE = """ & NVL(rsDetail!药品价格) & """"
                strDetail = strDetail & vbCrLf & "QUANTITY = """ & NVL(rsDetail!数量) & """"
                strDetail = strDetail & vbCrLf & "COSTS = """ & NVL(rsDetail!费用) & """"
                strDetail = strDetail & vbCrLf & "PAYMENTS = """ & NVL(rsDetail!实付费用) & """"
                strDetail = strDetail & vbCrLf & "DOSAGE = """ & NVL(rsDetail!药品剂量) & """"
                strDetail = strDetail & vbCrLf & "DOSAGE_UNITS = """ & SpecialChar(rsDetail!剂量单位) & """"
                strDetail = strDetail & vbCrLf & "ADMINISTRATION = """ & SpecialChar(rsDetail!用法) & """"
                strDetail = strDetail & vbCrLf & "FREQUENCY = """ & SpecialChar(rsDetail!执行频次) & """"
                strDetail = strDetail & vbCrLf & ">"
                strDetail = strDetail & vbCrLf & "</" & STR_DETAIL & ">"
                rsDetail.MoveNext
            Loop While Not rsDetail.EOF
            
            strBill = strBill & strDetail
            strBill = strBill & "</" & STR_BILL & ">"
            
            '拆分不同库房ID和病人ID的单据明细
            If lng库房ID = NVL(!发药药局, 0) And lng病人ID = NVL(!就诊卡号, 0) Then
                strXML = strXML & strBill & vbCrLf
            Else
                strXML = STR_ROOT_BEGIN & vbCrLf & strXML & STR_ROOT_END
                ReDim Preserve varReturn(UBound(varReturn) + 1)
                varReturn(UBound(varReturn)) = strXML
Call OutputLog(strXML)
                strXML = strBill & vbCrLf
            End If
            
            lng库房ID = NVL(!发药药局, 0)
            lng病人ID = NVL(!就诊卡号, 0)
            
            .MoveNext
        Loop While Not .EOF
        
        strXML = STR_ROOT_BEGIN & vbCrLf & strXML & STR_ROOT_END
        ReDim Preserve varReturn(UBound(varReturn) + 1)
        varReturn(UBound(varReturn)) = strXML
        varXML = varReturn
        GetXML_RecipeDetailEx = True
Call OutputLog(strXML)

    End With
    
    Exit Function
    
errHandle:
    Set varXML = Nothing
End Function

'Public Sub OutPutLog(ByVal strOutput As String)
'    '用于编译后的用户环境调试，在其他调试或不方便查找问题时使用
'    '将程序执行的关键流程，数据输出到外部日志文件，以此方便查找问题
'    '注意：在需要调试时手工创建指定的日志文件，编译环境时放到导航台程序所在目录，源代码环境时放到工程文件所在目录
'    '注意：如果不需要调试了要及时删除日志文件，否则日志文件可能会逐步增大，特别是用户环境可能数据增长较快
'    '各系统可以指定不同的日志文件名
'    '日志内容自定义，参考格式：时间+程序内部过程/函数+业务流程/步骤+关键数据
'    '默认的处理都加上时间，如果不需要可以去掉
'    Dim objFile As New FileSystemObject
'    Dim objTarget As TextStream
'    Const STR_CONS_FILENAME As String = "zlDrugPacker.log"
'
'    Err = 0
'
'    On Error Resume Next
'
'    '检查文件是否存在
'    Set objTarget = objFile.OpenTextFile(App.Path & "\" & STR_CONS_FILENAME)
'
'    '如果不存在则不输出内容
'    If objTarget Is Nothing Then Exit Sub
'
''    If err <> 0 Then
''        '创建目标文件
''        Set objFile = CreateObject("Scripting.FileSystemObject")
''        Set objTarget = objFile.CreateTextFile(App.Path & "\" & STR_CONS_FILENAME, True)
''        objTarget.Close
''    End If
'
'    Err.Clear
'    On Error GoTo errHand
'
'    Open App.Path & "\" & STR_CONS_FILENAME For Append Shared As #1
'
'    Print #1, strOutput
'    Close #1
'
'    Exit Sub
'errHand:
'    Close #1
''    MsgBox err.Description, vbExclamation + vbOKOnly
'End Sub

Public Sub OutputLog(ByVal strOutput As String)
'功能：将参数内容写入特定的日记文件中
'参数：
'  strOutput：日记内容

    Const STR_LOG_FILENAME As String = "zlDrugPacker"       '日志文本名称
    Const INT_MAX_DAY As Integer = 7                        '日志保存天数

    Dim objTS As TextStream
    Dim objFolder As Folder
    Dim objFile As File
    Dim strDate As String, strFileName As String
    Dim blnExist As Boolean, blnAutoCreate As Boolean

    On Error GoTo hErr

    '读取注册表的参数
    blnAutoCreate = Val(GetSetting("ZLSOFT", "公共模块\自动发药机", "自动生成日志")) = 1

    If blnAutoCreate Then
        '自动生成日志文件
        
        strFileName = STR_LOG_FILENAME & Format(Date, "_yyyymmdd") & ".log"
    
        ''判断文件是否存在
        Set objFolder = mobjFSO.GetFolder(App.Path)
        For Each objFile In objFolder.Files
            If LCase(objFile.Name) Like LCase(strFileName) Then
                blnExist = True
                Exit For
            End If
        Next
        
        Set objTS = mobjFSO.OpenTextFile(App.Path & "\" & strFileName, ForAppending, True)
        If blnExist = False Then
            '新创建的文件，强制加上时间戳
            strOutput = Now() & vbCrLf & strOutput
        End If
        objTS.WriteLine strOutput
        objTS.Close
        
        ''检查七天外的日志文件，并删除
        Set objFolder = mobjFSO.GetFolder(App.Path)
        For Each objFile In objFolder.Files
            If LCase(objFile.Name) Like LCase(STR_LOG_FILENAME) & "_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].log" Then
                strDate = Split(objFile.Name, "_")(1)
                strDate = Split(strDate, ".")(0)
                strDate = Left(strDate, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2)
                If Abs(Date - CDate(strDate)) >= INT_MAX_DAY Then
                    On Error Resume Next
                    objFile.Delete True
                    On Error GoTo hErr
                End If
            End If
        Next
    
    Else
        '不自动生成日志文件
        
        strFileName = STR_LOG_FILENAME & ".log"
    
        ''添加存储日志方式
        Set objTS = mobjFSO.OpenTextFile(App.Path & "\" & strFileName, ForAppending)
        If objTS Is Nothing Then Exit Sub
        
        objTS.WriteLine strOutput
        objTS.Close
    End If
    
    Exit Sub
    
hErr:
End Sub

Private Function GetXML_RecipeDetailEx_CQFLQZYY(ByVal rsBill As ADODB.Recordset, ByVal rsDetail As ADODB.Recordset, ByRef varXML As Variant) As Boolean
    '功能：处理库房ID为0的情况，分处理库房ID与病人ID生成XML字符串
    '参数：
    '  rsBill：单据数据集；
    '  rsDetail：明细数据集；
    '  varXML：生成的XML字符串数组（实参）。
    '返回：True成功   False失败
    '适用接口：韦乐海茨CONSIS系统v4.3
    Const STR_BILL = "CONSIS_PRESC_MSTVW"
    Const STR_DETAIL = "CONSIS_PRESC_DTLVW"
    Dim strXML As String, strBill As String, strDetail As String
    Dim lng库房ID As Long, lng病人ID As Long
    Dim varReturn As Variant
    Dim strOutput As String           '用于日志输出
    Dim strOutPutExeStep As String    '执行步骤，用于输出日志方便查找问题
    
    strOutput = "调用函数：GetXML_RecipeDetailEx_CQFLQZYY"
    
    On Error GoTo errHandle
    
    varReturn = Array()
    
    strOutPutExeStep = "初始化变量"
    
    With rsBill
        If .RecordCount <= 0 Then
            strOutput = strOutput & vbCrLf & "无数据"
            strOutput = strOutput & vbCrLf & "执行失败：GetXML_RecipeDetailEx_CQFLQZYY"
            Call OutputLog(strOutput)
            
            Exit Function
        End If
       
        .MoveFirst
        
        strOutPutExeStep = "库房ID，病人ID初始赋值"
        
        lng库房ID = NVL(!发药药局, 0)
        lng病人ID = NVL(!就诊卡号, 0)
        
        strOutPutExeStep = "拼装XML：begin"
        
        Do
            If .EOF Then Exit Do
            '单据
            strBill = "<" & STR_BILL & ">"
            
            strBill = strBill & vbCrLf & GetXMLFormat("PRESC_DATE", Format(!处方时间, "yyyy-MM-DDThh:mm:ss"))
            strBill = strBill & vbCrLf & GetXMLFormat("PRESC_NO", SpecialChar(!处方编号))
            strBill = strBill & vbCrLf & GetXMLFormat("DISPENSARY", NVL(!发药药局))
            strBill = strBill & vbCrLf & GetXMLFormat("PATIENT_ID", NVL(!就诊卡号))
            strBill = strBill & vbCrLf & GetXMLFormat("PATIENT_NAME", SpecialChar(!患者姓名))
            
            strOutPutExeStep = "拼装处方单XML：1"
            
            strBill = strBill & vbCrLf & GetXMLFormat("INVOICE_NO", "")
            strBill = strBill & vbCrLf & GetXMLFormat("PATIENT_TYPE", NVL(!患者类型))
            strBill = strBill & vbCrLf & GetXMLFormat("DATE_OF_BIRTH", Format(NVL(!患者出生日期), "yyyy-MM-DDThh:mm:ss"))
            strBill = strBill & vbCrLf & GetXMLFormat("SEX", SpecialChar(!患者性别))
            strBill = strBill & vbCrLf & GetXMLFormat("PRESC_IDENTITY", SpecialChar(!患者身份))
            
            strOutPutExeStep = "拼装处方单XML：2"
            
            strBill = strBill & vbCrLf & GetXMLFormat("CHARGE_TYPE", SpecialChar(!医保类型))
            strBill = strBill & vbCrLf & GetXMLFormat("PRESC_ATTR", "")
            strBill = strBill & vbCrLf & GetXMLFormat("PRESC_INFO", "")
            strBill = strBill & vbCrLf & GetXMLFormat("RCPT_INFO", GetRCPT_INFO(NVL(!处方编号)))
            strBill = strBill & vbCrLf & GetXMLFormat("RCPT_REMARK", "")
            
            strOutPutExeStep = "拼装处方单XML：3"
            
            strBill = strBill & vbCrLf & GetXMLFormat("REPETITION", "1")
            strBill = strBill & vbCrLf & GetXMLFormat("COSTS", NVL(!费用))
            strBill = strBill & vbCrLf & GetXMLFormat("PAYMENTS", NVL(!实付费用))
            strBill = strBill & vbCrLf & GetXMLFormat("ORDERED_BY", NVL(!开单科室编码))
            strBill = strBill & vbCrLf & GetXMLFormat("ORDERED_BY_NAME", NVL(!开单科室名称))
            
            strOutPutExeStep = "拼装处方单XML：4"
            
            strBill = strBill & vbCrLf & GetXMLFormat("PRESCRIBED_BY", SpecialChar(!开方医生))
            strBill = strBill & vbCrLf & GetXMLFormat("ENTERED_BY", SpecialChar(!录方人))
            strBill = strBill & vbCrLf & GetXMLFormat("DISPENSE_PRI", NVL(!配药优先级))
            
            strOutPutExeStep = "拼装处方单XML：5"
            
            '过滤明细记录，确保与单据对应
            strDetail = ""
            rsDetail.Filter = "no='" & !处方编号 & "' and 单据=" & NVL(!单据) & " and 库房id=" & NVL(!发药药局) & " and 病人id=" & NVL(!就诊卡号)
            rsDetail.Sort = "序号"
            
            strOutPutExeStep = "过滤明细记录"
            
            Do
                If rsDetail.EOF Then Exit Do
                '明细
                strDetail = strDetail & vbCrLf & "<" & STR_DETAIL & ">"
                
                strDetail = strDetail & vbCrLf & GetXMLFormat("PRESC_NO", NVL(rsDetail!no))
                strDetail = strDetail & vbCrLf & GetXMLFormat("ITEM_NO", NVL(rsDetail!序号))
                strDetail = strDetail & vbCrLf & GetXMLFormat("ADVICE_CODE", NVL(rsDetail!医嘱序号))
                strDetail = strDetail & vbCrLf & GetXMLFormat("DRUG_CODE", SpecialChar(rsDetail!药品编码))
                strDetail = strDetail & vbCrLf & GetXMLFormat("DRUG_SPEC", SpecialChar(rsDetail!药品规格))
                
                strOutPutExeStep = "拼装处方明细XML：1"
                
                strDetail = strDetail & vbCrLf & GetXMLFormat("DRUG_NAME", SpecialChar(rsDetail!药品名称))
                strDetail = strDetail & vbCrLf & GetXMLFormat("FIRM_ID", SpecialChar(rsDetail!生产厂家编码))
                strDetail = strDetail & vbCrLf & GetXMLFormat("FIRM_NAME", SpecialChar(rsDetail!生产厂家名称))
                strDetail = strDetail & vbCrLf & GetXMLFormat("PACKAGE_SPEC", SpecialChar(rsDetail!药品包装规格))
                strDetail = strDetail & vbCrLf & GetXMLFormat("PACKAGE_UNITS", SpecialChar(rsDetail!药品单位))
                
                strOutPutExeStep = "拼装处方明细XML：2"
                
                strDetail = strDetail & vbCrLf & GetXMLFormat("QUANTITY", NVL(rsDetail!数量))
                strDetail = strDetail & vbCrLf & GetXMLFormat("UNIT", SpecialChar(rsDetail!药品单位))
                strDetail = strDetail & vbCrLf & GetXMLFormat("COSTS", NVL(rsDetail!费用))
                strDetail = strDetail & vbCrLf & GetXMLFormat("PAYMENTS", NVL(rsDetail!实付费用))
                strDetail = strDetail & vbCrLf & GetXMLFormat("DOSAGE", NVL(rsDetail!药品剂量))
                
                strOutPutExeStep = "拼装处方明细XML：3"
                
                strDetail = strDetail & vbCrLf & GetXMLFormat("DOSAGE_UNITS", SpecialChar(rsDetail!剂量单位))
                strDetail = strDetail & vbCrLf & GetXMLFormat("ADMINISTRATION", SpecialChar(rsDetail!用法))
                strDetail = strDetail & vbCrLf & GetXMLFormat("FREQUENCY ", SpecialChar(rsDetail!执行频次))
                strDetail = strDetail & vbCrLf & GetXMLFormat("Additionusage", "")
                strDetail = strDetail & vbCrLf & GetXMLFormat("Rcpt_remark", "")
                
                strOutPutExeStep = "拼装处方明细XML：4"
                
                strDetail = strDetail & vbCrLf & "</" & STR_DETAIL & ">"
                rsDetail.MoveNext
            Loop While Not rsDetail.EOF
            
            strBill = strBill & strDetail
            strBill = strBill & "</" & STR_BILL & ">"
            
            '拆分不同库房ID和病人ID的单据明细
            If lng库房ID = NVL(!发药药局, 0) And lng病人ID = NVL(!就诊卡号, 0) Then
                strXML = strXML & strBill & vbCrLf
                strOutPutExeStep = "拆分不同库房ID和病人ID的单据明细"
            Else
                ReDim Preserve varReturn(UBound(varReturn) + 1)
                varReturn(UBound(varReturn)) = strXML
                strXML = strBill & vbCrLf
                
                strOutPutExeStep = "按不同库房和病人分别装入数组中"
            End If
            
            lng库房ID = NVL(!发药药局, 0)
            lng病人ID = NVL(!就诊卡号, 0)
            
            strOutPutExeStep = "库房ID，病人ID赋值"
            
            .MoveNext
        Loop While Not .EOF
        
        ReDim Preserve varReturn(UBound(varReturn) + 1)
        varReturn(UBound(varReturn)) = strXML
        varXML = varReturn
        
        strOutPutExeStep = "拼装XML：end"
    End With
    
    GetXML_RecipeDetailEx_CQFLQZYY = True
    
    strOutput = strOutput & vbCrLf & "执行成功：GetXML_RecipeDetailEx_CQFLQZYY"
    Call OutputLog(strOutput)
            
    Exit Function
    
errHandle:
    Set varXML = Nothing
    
    strOutput = strOutput & vbCrLf & "发生异常错误"
    strOutput = strOutput & vbCrLf & "执行失败：GetXML_RecipeDetail_CQFLQZYY"
    strOutput = strOutput & vbCrLf & "最后步骤：" & strOutPutExeStep
    Call OutputLog(strOutput)
End Function


Public Function GetXML_RecipeList(ByVal lngStockID As Long, ByVal strNO As String) As Variant
'将处方单组织成指定的XML格式
'适用接口：韦乐海茨CONSIS系统v2.2
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim strTitle As String
    Dim arrXML As Variant, arrTmp As Variant
    Dim i As Integer
    
    On Error GoTo errHandle
    
    mstrSQL = "Select 填制日期,No From 药品收发记录 Where 库房id=[1]"
    
    If InStr(1, strNO, "|") < 1 Then
        mstrSQL = mstrSQL & " And 单据=[2] And NO=[3]"
    Else
        mstrSQL = mstrSQL & " And ("
        arrTmp = Split(strNO, "|")
        For i = 0 To UBound(arrTmp)
            If i = UBound(arrTmp) Then
                mstrSQL = mstrSQL & "(单据=" & Split(arrTmp(i), ",")(0) & " And NO='" & Split(arrTmp(i), ",")(1) & "')"
            Else
                mstrSQL = mstrSQL & "(单据=" & Split(arrTmp(i), ",")(0) & " And NO='" & Split(arrTmp(i), ",")(1) & "') or "
            End If
        Next
        mstrSQL = mstrSQL & ")"
    End If
    mstrSQL = mstrSQL & " and (记录状态=1 or mod(记录状态,3)=1) "
    
    If InStr(1, strNO, "|") < 1 Then
        If gintMode = 0 Then
            Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_RecipeList", lngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
        Else
            Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_RecipeList", lngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
        End If
    Else
        If gintMode = 0 Then
            Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_RecipeList", lngStockID)
        Else
            Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_RecipeList", lngStockID)
        End If
    End If
    
    strXML = ""
    arrXML = Array()
    
    With rsTemp
        If .RecordCount > 0 Then
            strTitle = "<ROOT>"
            
            Do While Not .EOF
                strDrug = "<CONSIS_PRESC_MSTVW"
                strDrug = strDrug & vbCrLf & "PRESC_DATE = """ & Format(!填制日期, "yyyy-MM-DDThh:mm:ss") & """"
                strDrug = strDrug & vbCrLf & "PRESC_NO = """ & NVL(!no) & """"
                strDrug = strDrug & vbCrLf & ">"
                strDrug = strDrug & vbCrLf & "</CONSIS_PRESC_MSTVW>"
                
                strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
                rsTemp.MoveNext
                If .EOF Then
                    strXML = strXML & vbCrLf & "</ROOT>"
                    
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                    Call OutputLog(strXML)
                End If
            Loop
        End If
    End With
    
    GetXML_RecipeList = arrXML
    Exit Function
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
End Function

Public Function GetXML_RecipeList_CQFLQZYY(ByVal lngStockID As Long, ByVal strNO As String, ByRef strOutput As String) As Variant
    '将处方单组织成指定的XML格式
    '适用接口：韦乐海茨CONSIS系统v4.3
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim arrXML As Variant, arrTmp As Variant
    Dim i As Integer
    Dim strOutPutExeStep As String    '执行步骤，用于输出日志方便查找问题
    
'    Presc_date  处方时间    Datetime    N
'    Presc_no    处方编号    Nvarchar(200)   N
'    Invoice_no  发票编号    Nvarchar(200)   Y
'    DISPENSARY  发药药局编号
    
    On Error GoTo errHandle
    
    strOutput = strOutput & vbCrLf & "调用函数：GetXML_RecipeList_CQFLQZYY"
   
    mstrSQL = "Select 填制日期,No From 未发药品记录 Where 库房id=[1]"
    
    If InStr(1, strNO, "|") < 1 Then
        mstrSQL = mstrSQL & " And 单据=[2] And NO=[3]"
    Else
        mstrSQL = mstrSQL & " And ("
        arrTmp = Split(strNO, "|")
        For i = 0 To UBound(arrTmp)
            If i = UBound(arrTmp) Then
                mstrSQL = mstrSQL & "(单据=" & Split(arrTmp(i), ",")(0) & " And NO='" & Split(arrTmp(i), ",")(1) & "')"
            Else
                mstrSQL = mstrSQL & "(单据=" & Split(arrTmp(i), ",")(0) & " And NO='" & Split(arrTmp(i), ",")(1) & "') or "
            End If
        Next
        mstrSQL = mstrSQL & ")"
    End If
    
    strOutPutExeStep = "拼凑SQL"
    
    If InStr(1, strNO, "|") < 1 Then
        If gintMode = 0 Then
            Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_RecipeList_CQFLZYY", lngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
        Else
            Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_RecipeList_CQFLZYY", lngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
        End If
    Else
        If gintMode = 0 Then
            Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_RecipeList_CQFLZYY", lngStockID)
        Else
            Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_RecipeList_CQFLZYY", lngStockID)
        End If
    End If
    
    strOutPutExeStep = "执行SQL成功"
    
    strXML = ""
    arrXML = Array()
    
    With rsTemp
        strOutPutExeStep = "拼装XML：begin"
        
        If .RecordCount > 0 Then
            Do While Not .EOF
                strDrug = "<CONSIS_PRESC_MSTVW>"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRESC_DATE", Format(!填制日期, "yyyy-MM-DDThh:mm:ss"))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRESC_NO", NVL(!no))
                strDrug = strDrug & vbCrLf & GetXMLFormat("INVOICE_NO", "")
                strDrug = strDrug & vbCrLf & GetXMLFormat("DISPENSARY", lngStockID)
                
                strOutPutExeStep = "拼装XML：1"
                
                strDrug = strDrug & vbCrLf & "</CONSIS_PRESC_MSTVW>"
                
                strXML = IIf(strXML = "", "", strXML & vbCrLf) & strDrug
                
                rsTemp.MoveNext
                
                If .EOF Then
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                End If
            Loop
            
            strOutPutExeStep = "拼装XML：end"
        End If
    End With
    
    GetXML_RecipeList_CQFLQZYY = arrXML
    
    strOutput = strOutput & vbCrLf & "执行成功：GetXML_RecipeList_CQFLQZYY"
    
    Exit Function
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
    
    strOutput = strOutput & vbCrLf & "发生异常错误"
    strOutput = strOutput & vbCrLf & "最后步骤：" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "相关SQL" & vbCrLf & mstrSQL
    strOutput = strOutput & vbCrLf & "执行失败：GetXML_RecipeList_CQFLQZYY"
    Call OutputLog(strOutput)
End Function

Public Function IsRegisterStock(ByVal lngStockID As Long, ByVal strStockIDs As String) As Boolean
'功能：检查是否为注册药房
    Dim i As Integer
    Dim arrID As Variant
    
    If Val(strStockIDs) = 0 Or lngStockID = 0 Then Exit Function
    
    arrID = Split(strStockIDs, ";")
    For i = LBound(arrID) To UBound(arrID)
        If Val(arrID(i)) = lngStockID Then
            IsRegisterStock = True
            Exit For
        End If
    Next
End Function

Public Function GetXML_RecipeReturn_CQFLQZYY(ByVal strReturnRecipt As String, ByVal strStockIDs As String, ByRef strOutput As String) As Variant
'将处方单组织成指定的XML格式
'退费处方信息
'strReturnRecipt：退费处方信息，格式：NO,药房id|NO,药房id
'适用接口：韦乐海茨CONSIS系统v4.3
    Dim strXML As String
    Dim arrXML As Variant
    Dim arrRecipt
    Dim n As Integer
    Dim strOutPutExeStep As String    '执行步骤，用于输出日志方便查找问题
    
'    Presc_date  处方时间    Datetime    N
'    Presc_no    处方编号    Nvarchar(200)   N
'    DISPENSARY  发药药局编号

    strOutput = strOutput & vbCrLf & "调用函数：GetXML_RecipeReturn_CQFLQZYY"
    
    On Error GoTo errHandle
    
    arrRecipt = Split(strReturnRecipt, "|")
    arrXML = Array()
    
    strOutPutExeStep = "拼装XML：begin"
    
    For n = 0 To UBound(arrRecipt)
        '注册的药房才提交数据
        If IsRegisterStock(Val(Split(arrRecipt(n), ",")(1)), strStockIDs) Then
            strXML = IIf(strXML = "", "", strXML & vbCrLf) & "<CONSIS_PRESC_MSTVW>"
                        
            strXML = strXML & vbCrLf & GetXMLFormat("PRESC_DATE", Format(CStr(Now), "yyyy-MM-DDThh:mm:ss"))
            strXML = strXML & vbCrLf & GetXMLFormat("PRESC_NO", Split(arrRecipt(n), ",")(0))
            strXML = strXML & vbCrLf & GetXMLFormat("INVOICE_NO", "")
            strXML = strXML & vbCrLf & GetXMLFormat("DISPENSARY", Split(arrRecipt(n), ",")(1))
            
            strXML = strXML & vbCrLf & "</CONSIS_PRESC_MSTVW>"
        End If
    Next
    
    If strXML <> "" Then
        ReDim Preserve arrXML(UBound(arrXML) + 1)
        arrXML(UBound(arrXML)) = strXML
    End If
    
    strOutPutExeStep = "拼装XML：end"
    
    GetXML_RecipeReturn_CQFLQZYY = arrXML
    
    strOutput = strOutput & vbCrLf & "执行成功：GetXML_RecipeReturn_CQFLQZYY"
    
    Exit Function
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
    
    strOutput = strOutput & vbCrLf & "发生异常错误"
    strOutput = strOutput & vbCrLf & "最后步骤：" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "执行失败：GetXML_RecipeReturn_CQFLQZYY"
    Call OutputLog(strOutput)
End Function


Public Function GetXML_Stock(ByVal lngStockID As Long) As Variant
'将药品库存信息组织成指定的XML格式
'适用接口：韦乐海茨CONSIS系统v2.2
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim strTitle As String
    Dim arrXML As Variant
    
    On Error GoTo errHandle
    mstrSQL = "Select a.id 药品编号,c.库房id 发药药局,sum(c.实际数量/e.门诊包装) 药品数量,d.库房货位 药品货位 " & vbNewLine & _
              "From 收费项目目录 a, 药品库存 c, 药品储备限额 d,药品规格 e " & vbNewLine & _
              "Where a.Id = c.药品id And e.药品id=c.药品id And d.库房id(+) = c.库房id And d.药品id(+) = c.药品id And c.库房id=[1] " & vbNewLine & _
              "Group By a.id, c.库房id, d.库房货位 " & vbNewLine & _
              "Having Sum(c.实际数量/e.门诊包装)<>0 "
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_Stock", lngStockID)
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_Stock", lngStockID)
    End If
    strXML = ""
    arrXML = Array()
    
    With rsTemp
        If .RecordCount > 0 Then
            strTitle = "<ROOT>"
            
            Do While Not .EOF
                strDrug = "<CONSIS_PHC_STORAGEVW"
                strDrug = strDrug & vbCrLf & "DRUG_CODE = """ & SpecialChar(!药品编号) & """"
                strDrug = strDrug & vbCrLf & "DISPENSARY = """ & NVL(!发药药局) & """"
                strDrug = strDrug & vbCrLf & "DRUG_QUANTITY = """ & NVL(!药品数量) & """"
                strDrug = strDrug & vbCrLf & "LOCATIONINFO = """ & SpecialChar(!药品货位) & """"
                strDrug = strDrug & vbCrLf & ">"
                strDrug = strDrug & vbCrLf & "</CONSIS_PHC_STORAGEVW>"

'该业务功能可以不用4K限制
                strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
                
'                If Len(strXML & strDrug) > 3900 Then
'                    '将以前的添加到数组
'                    strXML = strXML & vbCrLf & "</ROOT>"
'                    ReDim Preserve arrXML(UBound(arrXML) + 1)
'                    arrXML(UBound(arrXML)) = strXML
'
'                    '重新拼凑新的XML
'                    strXML = strTitle & vbCrLf & strDrug
'                Else
'                    strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
'                End If
                
                rsTemp.MoveNext
                If .EOF Then
                    strXML = strXML & vbCrLf & "</ROOT>"
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                End If
            Loop
        End If
    End With
    
    GetXML_Stock = arrXML
    Exit Function
    
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
End Function

Public Function GetXML_Dept(ByVal strProperty As String, Optional ByRef strLog As String) As Variant
'功能：获取ZLHIS的部门数据，并转换成“韦乐海茨”接口要求的XML格式
'参数：
'  strLog：日志变量
'返回：XML字符串数组

    Dim rsTemp As ADODB.Recordset
    Dim arrXML As Variant
    Dim objXML As clsXML
    
strLog = strLog & "获取部门数据开始！" & vbCrLf
    
    mstrSQL = "Select Distinct a.Id, a.名称, b.服务对象 " & vbNewLine & _
              "From 部门表 A, 部门性质说明 B " & vbNewLine & _
              "Where a.Id = b.部门id And Trunc(Nvl(a.撤档时间, To_Date('3000-1-1', 'YYYY-MM-DD'))) = To_Date('3000-1-1', 'YYYY-MM-DD') " & _
              "    And b.服务对象 <> 0 "
    
    If strProperty <> "" Then
        strProperty = "," & strProperty & ","
        mstrSQL = mstrSQL & " And Instr([1], ',' || b.工作性质 || ',') > 0 "
    End If
    mstrSQL = mstrSQL & vbNewLine & _
              "Order By a.ID "
              
    On Error GoTo hErr
    
strLog = strLog & "SQL：" & mstrSQL & vbCrLf
    
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_Dept", strProperty)
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_Dept", strProperty)
    End If
    
strLog = strLog & "准备组装XML！" & vbCrLf
    
    With rsTemp
        arrXML = Array()
        Set objXML = New clsXML
        Do While .EOF = False
            objXML.ClearXmlText
            Call objXML.AppendNode("CONSIS_BASIC_DEPTVW")
                Call objXML.AppendData("DEPTCODE", NVL(!ID))
                Call objXML.AppendData("DEPTNAME", NVL(!名称))
                Call objXML.AppendData("OUTP_OR_INP", NVL(!服务对象))
            Call objXML.AppendNode("CONSIS_BASIC_DEPTVW", True)
            
'strLog = strLog & "XML：" & objXML.XmlText
            
            ReDim Preserve arrXML(UBound(arrXML) + 1)
            arrXML(UBound(arrXML)) = objXML.XmlText
            
'strLog = strLog & "完成！" & vbCrLf
            
            .MoveNext
        Loop
        .Close
    End With
    
strLog = strLog & "组装XML完成！" & vbCrLf
    
    GetXML_Dept = arrXML
    
    Exit Function
    
hErr:
strLog = strLog & "获取部门数据异常！" & Err.Description & vbCrLf
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
End Function

Public Function GetXML_Stock_CQFLQZYY(ByVal lngStockID As Long, ByRef strOutput As String) As Variant
'将药品库存信息组织成指定的XML格式
'适用接口：韦乐海茨CONSIS系统v4.3
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim arrXML As Variant
    Dim strOutPutExeStep As String    '执行步骤，用于输出日志方便查找问题
    
'    字段名  类型    说明    NULL
'    Dispensary  Nvarchar(40)    发药药局    N
'    Drug_code   Nvarchar(40)    药品编号    N
'    Locationinfo    Nvarchar(200)   货位信息    N
'    Batchid Nvarchar(200)   药品批次    Y
'    Batchno Nvarchar(200)   药品批号    Y
'    Producedate Datetime    生产日期    Y
'    Disableddate    Datetime    失效日期    Y
'    Quantity    Numeric(20,6)   药品货位库存数量    N
    
    On Error GoTo errHandle
    
    strOutput = strOutput & vbCrLf & "调用函数：GetXML_Stock_CQFLQZYY"

    mstrSQL = "Select a.编码 药品编号, c.库房id 发药药局, c.实际数量 / e.门诊包装 As 药品数量, Nvl(d.库房货位, '无') As 药品货位, Nvl(c.批次, 0) As 批次, c.上次批号, c.效期, c.上次生产日期" & vbNewLine & _
        " From 收费项目目录 A, 药品库存 C, 药品储备限额 D, 药品规格 E" & vbNewLine & _
        " Where a.Id = c.药品id And e.药品id = c.药品id And d.库房id(+) = c.库房id And d.药品id(+) = c.药品id And c.库房id = [1] " & vbNewLine & _
        " Order By a.Id "
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_Stock_CQFLQZYY", lngStockID)
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_Stock_CQFLQZYY", lngStockID)
    End If
    
    strOutPutExeStep = "执行SQL成功"
    
    strXML = ""
    arrXML = Array()
    
    With rsTemp
        strOutPutExeStep = "拼装XML：begin"
        
        If .RecordCount > 0 Then
            Do While Not .EOF
                strDrug = "<CONSIS_BASIC_LOCATIONVW>"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("DISPENSARY", NVL(!发药药局))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DRUG_CODE", SpecialChar(!药品编号))
                strDrug = strDrug & vbCrLf & GetXMLFormat("LOCATIONINFO", SpecialChar(!药品货位))
                strDrug = strDrug & vbCrLf & GetXMLFormat("BATCHID", NVL(!批次))
                strDrug = strDrug & vbCrLf & GetXMLFormat("BATCHNO", SpecialChar(NVL(!上次批号)))
                
                strOutPutExeStep = "拼装XML：1"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRODUCEDATE", Format(NVL(!上次生产日期), "yyyy-MM-DDThh:mm:ss"))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DISABLEDDATE", Format(NVL(!效期), "yyyy-MM-DDThh:mm:ss"))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DRUG_QUANTITY", NVL(!药品数量))
                
                strOutPutExeStep = "拼装XML：2"
                
                strDrug = strDrug & vbCrLf & "</CONSIS_BASIC_LOCATIONVW>"

                strXML = IIf(strXML = "", "", strXML & vbCrLf) & strDrug
                
                rsTemp.MoveNext
                
                If .EOF Then
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                End If
            Loop
            
            strOutPutExeStep = "拼装XML：end"
        End If
    End With
    
    GetXML_Stock_CQFLQZYY = arrXML
    
    strOutput = strOutput & vbCrLf & "执行成功：GetXML_Stock_CQFLQZYY"
 
    Exit Function
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If

    strOutput = strOutput & vbCrLf & "发生异常错误"
    strOutput = strOutput & vbCrLf & "最后步骤：" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "相关SQL" & vbCrLf & mstrSQL
    strOutput = strOutput & vbCrLf & "执行失败：GetXML_Stock_CQFLQZYY"
    Call OutputLog(strOutput)
End Function


Public Function GetXMLFormat(ByVal strNode As String, ByVal strText As String, Optional ByVal blnNodeUpper As Boolean = True) As String
    '传入节点和数据，组合成XML格式
    '格式：<NODE>Text</NODE>
    strNode = Replace(strNode, "<", "")
    strNode = Replace(strNode, "</", "")
    strNode = Replace(strNode, ">", "")
    If blnNodeUpper = True Then
        GetXMLFormat = "<" & UCase(strNode) & ">" & strText & "</" & UCase(strNode) & ">"
    Else
        GetXMLFormat = "<" & strNode & ">" & strText & "</" & strNode & ">"
    End If
End Function

Public Function SetSendWin(ByVal lngStockID As Long, ByVal strNO As String, ByVal intOpr As Integer) As Boolean
'设置HIS中指定处方的发药窗口
    Dim i As Integer
    Dim arrTmp As Variant
    Dim rsTemp As Recordset
    
    On Error GoTo errHandle
    mstrSQL = "Select 名称 From 发药窗口 Where 药房id=[1] And 编码=[2]"
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "SetSendWin", lngStockID, CStr(intOpr))
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "SetSendWin", lngStockID, CStr(intOpr))
    End If
    
    If Not rsTemp.EOF Then
        arrTmp = Split(strNO, "|")
        For i = 0 To UBound(Split(strNO, "|"))
            mstrSQL = "Zl_未发药品记录_分配发药窗口("
            mstrSQL = mstrSQL & "'" & Split(arrTmp(i), ",")(1) & "',"
            mstrSQL = mstrSQL & Split(arrTmp(i), ",")(0) & ","
            mstrSQL = mstrSQL & lngStockID & ","
            mstrSQL = mstrSQL & "'" & rsTemp!名称 & "')"
            
            Call OutputLog(mstrSQL)
            If gintMode = 0 Then
                Call gobjComLib.zlDatabase.ExecuteProcedure(mstrSQL, "SetSendWin")
            Else
                Call mdlDrugPacker.ExecuteProcedure(mstrSQL, "SetSendWin")
            End If
        Next
        SetSendWin = True
    Else
        If gblnShowMsg Then
            MsgBox "没有找到编码为【" & intOpr & "】的窗口，请检查“发药窗口管理”模块！", vbCritical, GSTR_MESSAGE
        Else
            Call OutputLog("没有找到编码为【" & intOpr & "】的窗口，请检查“发药窗口管理”模块！")
        End If
    End If
    
    Exit Function
    
errHandle:
    If gblnShowMsg Then
        If gintMode = 0 Then
            If gobjComLib.ErrCenter() = 1 Then Resume
            Call gobjComLib.SaveErrLog
        Else
            MsgBox Err.Description, vbInformation, GSTR_SYSNAME
        End If
    End If
    Call OutputLog("SetSendWin异常！ " & Err.Description)
End Function


Public Function GetLocalIP() As String
'取本机IP
    Dim Ret As Long, Tel As Long
    Dim bBytes() As Byte
    Dim TempList() As String
    Dim TempIP As String
    Dim Tempi As Long
    Dim Listing As MIB_IPADDRTABLE
    Dim L3 As String
    
    
    On Error GoTo EndRow
        GetIpAddrTable ByVal 0&, Ret, True
    
    
        If Ret <= 0 Then Exit Function
        ReDim bBytes(0 To Ret - 1) As Byte
        ReDim TempList(0 To Ret - 1) As String
        
        'retrieve the data
        GetIpAddrTable bBytes(0), Ret, False
          
        'Get the first 4 bytes to get the entry's.. ip installed
        CopyMemory Listing.dEntrys, bBytes(0), 4
        
        For Tel = 0 To Listing.dEntrys - 1
            'Copy whole structure to Listing..
            CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
            TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr)
        Next Tel
        'Sort Out The IP For WAN
        TempIP = TempList(0)
        For Tempi = 0 To Listing.dEntrys - 1
            L3 = Left(TempList(Tempi), 3)
            If L3 <> "169" And L3 <> "127" And L3 <> "192" Then
                TempIP = TempList(Tempi)
            End If
        Next Tempi
        GetLocalIP = TempIP 'Return The TempIP
    Exit Function
EndRow:
    GetLocalIP = ""
End Function

Private Function ConvertAddressToString(longAddr As Long) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function


Private Function GetRCPT_INFO(ByVal strNO As String) As String
'功能：获取诊断信息
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "Select MAX(DECODE(Id,1,诊断描述,''))||';'||MAX(DECODE(Id,2,诊断描述,'')) as 诊断 " & vbNewLine & _
             "From ( " & vbNewLine & _
             "      Select Rownum As Id,诊断描述 " & vbNewLine & _
             "      From (Select 诊断描述||decode(是否疑诊,1,'?','') 诊断描述 " & vbNewLine & _
             "            From 病人诊断记录 " & vbNewLine & _
             "            Where 病人id=(Select distinct 病人id " & vbNewLine & _
             "                          From ( Select a.病人id From 门诊费用记录 a Left Join 病人医嘱记录 b On a.医嘱序号=b.Id " & vbNewLine & _
             "                                 Where a.No=[1] And 记录性质=1 ) ) " & vbNewLine & _
             "              And 主页id=(Select distinct Case When 主页id Is Null Then (Select Id From 病人挂号记录 Where No=c.挂号单) Else 主页Id End As 主页id " & vbNewLine & _
             "                          From ( Select null 主页id, b.挂号单 From 门诊费用记录 a Left Join 病人医嘱记录 b On a.医嘱序号=b.Id " & vbNewLine & _
             "                                 Where a.No=[1] And 记录性质=1 ) c ) " & vbNewLine & _
             "union all " & vbNewLine & _
             "Select a.摘要 As 诊断描述 From 病人挂号记录 a " & vbNewLine & _
             "Where No= (Select distinct Case When b.挂号单 Is Null Then ' ' Else b.挂号单 End As No " & vbNewLine & _
             "           From 门诊费用记录 a Left Join 病人医嘱记录 b On a.医嘱序号 = b.Id " & vbNewLine & _
             "           Where a.No = [1] And 记录性质 = 1 ) ) ) "
    On Error GoTo errHandle
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "获取诊断信息", strNO)
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(strSQL, "获取诊断信息", strNO)
    End If
    
    If Not rsTemp.EOF Then
        GetRCPT_INFO = IIf(Trim(NVL(rsTemp!诊断)) = ";", """""", """" & Trim(NVL(rsTemp!诊断)) & """")
    Else
        GetRCPT_INFO = """"""
    End If
    rsTemp.Close
    Exit Function
    
errHandle:
    GetRCPT_INFO = """"""
End Function

Private Function SpecialChar(ByVal strVal As Variant) As String
'功能：特殊字符转换
'说明：
' < 转 &lt;
' > 转 &gt;
' & 转 &amp;
' ' 转 &apos;
' " 转 &quot;
    Dim strReturn As String
    
    If IsNull(strVal) Then
        strVal = ""
        GoTo errHandle
    End If
    If strVal = "" Then
        GoTo errHandle
    End If
    On Error GoTo errHandle
    strReturn = strVal
    strReturn = Replace(strReturn, "<", "&lt;")
    strReturn = Replace(strReturn, ">", "&gt;")
    strReturn = Replace(strReturn, "&", "&amp;")
    strReturn = Replace(strReturn, "'", "&apos;")
    strReturn = Replace(strReturn, """", "&quot;")
    SpecialChar = strReturn
    Exit Function
    
errHandle:
    SpecialChar = strVal
End Function

Private Function GetStockID(ByVal strText As String) As Long
    '功能：获取XML文本中的药房ID
    Const STR_KEY = "DISPENSARY = "
    Const STR_NODES_BEGIN = "<DISPENSARY>"
    Const STR_NODES_END = "</DISPENSARY>"
    
    Dim lngStockID As Long
    Dim intStart As Integer
    Dim strTmp As String
    
    If strText = "" Then Exit Function
    
    intStart = InStr(strText, STR_KEY)
    If intStart > 0 Then
        lngStockID = Val(Mid(strText, intStart + Len(STR_KEY) + 1))
    Else
        strTmp = Mid(strText, 1, InStr(1, strText, STR_NODES_END) - 1)
        lngStockID = Val(Mid(strTmp, InStr(1, strTmp, STR_NODES_BEGIN) + Len(STR_NODES_BEGIN)))
    End If
    
    GetStockID = lngStockID
    
End Function

Public Function PackingWindow_DYEY(ByVal strNO As String, Optional ByRef strOut As String) As String
'功能：获取传入单据号的病人ID、发药药房、发药窗口、药品名称信息。主要是移动业务在使用
'参数：
'  strNO：单据信息，格式详见调用层说明
'  strOut（实参）：异常信息
'返回：病人ID、发药药房、发药窗口、药品名称信息
'XML格式：
'<OUTPUT>
'  <BRID>病人ID</BRID>
'  <ITEM>
'    <YFMC>药房名称</YFMC>
'    <YFCK>发药窗口</YFCK>
'    <YFMX>
'      <ITEM>
'        <MC>药品名称1</MC>
'      </ITEM>
'      <ITEM>
'        <MC>药品名称2</MC>
'      </ITEM>
'      <ITEM>
'        <MC>药品名称...</MC>
'      </ITEM>
'    </YFMX>
'  </ITEM>
'  <ITEM>
'    ...
'  </ITEM>
'</OUTPUT>

    Const STR_OUT_B As String = "<OUTPUT>", STR_OUT_E As String = "</OUTPUT>"
    Const STR_BRID_B As String = "<BRID>", STR_BRID_E As String = "</BRID>"
    Const STR_ITEM_B As String = "<ITEM>", STR_ITEM_E As String = "</ITEM>"
    Const STR_YFMC_B As String = "<YFMC>", STR_YFMC_E As String = "</YFMC>"
    Const STR_YFCK_B As String = "<YFCK>", STR_YFCK_E As String = "</YFCK>"
    Const STR_YFMX_B As String = "<YFMX>", STR_YFMX_E As String = "</YFMX>"
    Const STR_MC_B As String = "<MC>", STR_MC_E As String = "</MC>"

    Dim rsSQL As ADODB.Recordset
    Dim strSQL As String, strReturn As String, strDrugs As String
    Dim strStore As String, strWin As String
    Dim lngStoreID As Long
    
    On Error GoTo errHandle
    
    strSQL = "Select Distinct b.病人id, a.库房id, d.名称 As 药房名称, a.发药窗口, c.名称 As 药品名称 " & vbCr & _
             "From 药品收发记录 A, 门诊费用记录 B, 收费项目目录 C, 部门表 D, Table(f_Str2list2([1], '|', ',')) E " & vbCr & _
             "Where a.费用id = b.Id And a.药品id = c.Id And a.库房id = d.Id And a.单据 = e.C1 And a.No = e.C2 " & vbCr & _
             "Order By Nvl(b.病人id, 0) Desc, a.库房id, c.名称 "
    Set rsSQL = mdlDrugPacker.OpenSQLRecord(strSQL, "获取药房发药信息", strNO)
    
    With rsSQL
        If .EOF = False Then
            lngStoreID = NVL(!库房id, 0)
            strReturn = STR_OUT_B & vbCr & _
                        STR_BRID_B & NVL(!病人id) & STR_BRID_E & vbCr
            strWin = NVL(!发药窗口)
        End If
        Do While .EOF = False
            If lngStoreID = NVL(!库房id, 0) Then
                If strWin = "" Then strWin = NVL(!发药窗口)
                strDrugs = strDrugs & STR_ITEM_B & vbCr & STR_MC_B & NVL(!药品名称) & STR_MC_E & vbCr & STR_ITEM_E & vbCr
            Else
                strReturn = strReturn & _
                            strStore & _
                            STR_YFCK_B & strWin & STR_YFCK_E & vbCr & _
                            STR_YFMX_B & vbCr & strDrugs & STR_YFMX_E & vbCr & _
                            STR_ITEM_E & vbCr
                strDrugs = STR_ITEM_B & vbCr & STR_MC_B & NVL(!药品名称) & STR_MC_E & vbCr & STR_ITEM_E & vbCr
                strWin = NVL(!发药窗口)
            End If
            
            strStore = STR_ITEM_B & vbCr & STR_YFMC_B & NVL(!药房名称) & STR_YFMC_E & vbCr
            lngStoreID = NVL(!库房id, 0)
            .MoveNext
        Loop
        If .RecordCount > 0 Then
            strReturn = strReturn & _
                        strStore & _
                        STR_YFCK_B & strWin & STR_YFCK_E & vbCr & _
                        STR_YFMX_B & vbCr & strDrugs & STR_YFMX_E & vbCr & _
                        STR_ITEM_E & vbCr & _
                        STR_OUT_E
            strDrugs = ""
        End If
        
        .Close
        
    End With
    
    PackingWindow_DYEY = strReturn
    Exit Function
    
errHandle:
    strOut = strOut & vbCrLf & "返回“病人ID、发药药房、发药窗口、药品名称”信息失败"
    PackingWindow_DYEY = ""
End Function

Public Function HaveUpdateFlag() As Boolean
'功能：检查是否需要上传标志更新
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select Count(1) Rec From All_Tab_Columns Where Table_Name = [1] And Column_Name = [2] ANd Rownum < 2"
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "检查标志", "未发药品记录", "是否上传")
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(strSQL, "检查标志", "未发药品记录", "是否上传")
    End If
    If rsTemp!Rec > 0 Then
        HaveUpdateFlag = True
    End If
    rsTemp.Close
    Exit Function
    
errHandle:
    Call OutputLog("检查“是否上传”字段失败：" & Err.Description)
End Function

Public Function UpdateFlag(ByVal lngStockID As Long, ByVal strNO As String) As Boolean
    Dim strSQL As String
    Dim strTmp As String
    Dim arrNO As Variant, arrItem As Variant
    Dim l As Long
    Dim strRecipe As String
    
    On Error GoTo errHandle
    
    strTmp = "更新上传标志开始" & vbNewLine
    strTmp = strTmp & "参数1：" & lngStockID & "； 参数2：" & strNO & vbNewLine
    
    If Trim(strNO) <> "" Then
        
        If gblnUpdateFlag = False Then
            '无“是否上传”字段，无需更新上传标志
            UpdateFlag = True
            Exit Function
        Else
            strTmp = strTmp & "需要更新标志！" & vbNewLine
        End If
    
        arrNO = Split(strNO, "|")
        For l = LBound(arrNO) To UBound(arrNO)
            strRecipe = arrNO(l)
            If strRecipe <> "" Then
                strSQL = "Zl_未发药品记录_更新上传标志(" _
                       & lngStockID & "," _
                       & "'" & strRecipe & "')"
                If gintMode = 0 Then
                    Call gobjComLib.zlDatabase.ExecuteProcedure(strSQL, "更新标志")
                Else
                    Call mdlDrugPacker.ExecuteProcedure(strSQL, "更新标志")
                End If
                strTmp = strTmp & arrNO(l) & " 完成！" & vbNewLine
            End If
        Next
    Else
        strTmp = strTmp & "无单据更新！"
    End If
    UpdateFlag = True
    Call OutputLog(strTmp & vbCrLf & "更新上传标志完成")
    Exit Function
    
errHandle:
    strTmp = strTmp & vbNewLine & "更新上传标志异常:"
    Call OutputLog(strTmp & Err.Description)
End Function
