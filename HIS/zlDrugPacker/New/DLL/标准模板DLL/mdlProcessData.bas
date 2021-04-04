Attribute VB_Name = "mdlProcessData"
Option Explicit

Public Function ProcDrugInfo(ByVal strDrugType As String, ByVal objDevice As clsDevice) As ADODB.Recordset
'功能：获取HIS药品基本信息
'参数：
'  strDrugType：剂型串
'  objDevice：设备对象
'返回：已格式化的记录集

    Dim strSQL As String
    Dim rsData As ADODB.Recordset, rsUpload As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '读HIS数据
    Set rsData = mdlDefine.GetHisRecord_DrugInf(1, strDrugType)
    
    '格式化要上传的数据
    Set rsUpload = BuildDrugInfo(rsData, objDevice)
    
    If Not rsUpload Is Nothing Then
        Set ProcDrugInfo = rsUpload
    End If
    
    Exit Function

errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Function

Public Function ProcDrugStock(ByVal lngDeptID As Long, ByVal objDevice As clsDevice) As ADODB.Recordset
'功能：获取HIS药品库存信息
'参数：
'  lngDeptID：药房ID
'  objDevice：设备对象
'返回：已格式化的记录集
    
    Dim strSQL As String
    Dim rsData As ADODB.Recordset, rsUpload As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '读HIS药品库存数据
    Set rsData = mdlDefine.GetHisRecord_DrugStock(lngDeptID)
    
    '格式化要上传的数据
    Set rsUpload = BuildDrugStock(rsData, objDevice)
    
    If Not rsUpload Is Nothing Then
        Set ProcDrugStock = rsUpload
    End If
    
    Exit Function
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Function

Public Function SetUpload(ByVal bytType As Byte, ByVal varKey As Variant, ByVal lngModule As Long) As ADODB.Recordset
'功能：获取HIS相关上传信息
'参数：
'   bytType：
'       1: 门诊处方上传 (配药)
'       2: 门诊发药通知 (发药)
'       3: 住院药品医嘱上传 (配、发药)
'   varKey：
'       当bytType=1时，varKey格式：“单据|NO1,NO2...”；
'       当bytType=2时，varKey格式：“单据;库房ID;NO|单据;库房ID;NO...”；
'       当bytType=3时，varKey格式：药品收发ID1,药品收发ID2...
'  lngModule：HIS业务模块号
'返回：已格式化的记录集

    Dim strSQL As String
    Dim rsData As ADODB.Recordset, rsUpload As ADODB.Recordset
    'Dim arrBill As Variant
    'Dim i As Integer

    '读HIS数据
    Select Case bytType
    Case 1
        '门诊处方明细
        Set rsData = mdlDefine.GetHisRecord_ReceipInf(varKey)
        '格式化要上传的数据
        Set rsUpload = BuildReceipDetail(rsData, lngModule)
        
    Case 2
        '门诊发药通知
        Set rsData = mdlDefine.GetHisRecord_ReceipList(varKey)
        '格式化要上传的数据
        Set rsUpload = BuildReceipList(rsData, lngModule)
        
    Case 3
        '住院药品医嘱
        Set rsData = mdlDefine.GetHisRecord_AdviceInf(varKey)
        '格式化要上传的数据
        Set rsUpload = BuildReceipAdviceInf(rsData, lngModule)
        
    End Select
    
    If Not rsData Is Nothing Then
        Set SetUpload = rsUpload
    End If
End Function

Private Function BuildDrugInfo(ByVal rsDrugInfo As ADODB.Recordset, ByVal objDevice As clsDevice) As ADODB.Recordset
'功能：构建符合药品信息上传数据结构的记录集对象
'参数：
'  rsDrugInfo：HIS药品信息记录集对象
'  objDevice：设备对象

    Dim i As Integer
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    
    If rsDrugInfo Is Nothing Then Exit Function
    
    '初始化内存记录集对象
    With rsData
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "SN", adInteger, 10, adFldIsNullable
        .Fields.Append "Drug", adVarChar, 100, adFldIsNullable
        .Fields.Append "Content", adVarChar, 3000, adFldIsNullable
        .Open
    End With
    
    With rsDrugInfo
        If .State <> adStateOpen Then .Open
        i = 1
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
            '格式化成要上传的数据格式
            Select Case objDevice.LinkType
            Case enuLinkType.DB
                strTmp = ""
            
            Case enuLinkType.WEBServices
                strTmp = ""
                
            Case enuLinkType.Directory
                strTmp = ""
                
            End Select
            
            '存入内存记录集
            If strTmp <> "" Then
                rsData.AddNew
                rsData!SN = i
                rsData!Drug = !编码 & "；" & !通用名 & "；" & NVL(!规格)
                rsData!Content = strTmp
                rsData.Update
                i = i + 1
            End If
            
            .MoveNext
        Loop
        .Close
        
    End With
    Set BuildDrugInfo = rsData
    
End Function

Private Function BuildReceipDetail(ByVal rsVal As ADODB.Recordset, ByVal lngModule As Long) As ADODB.Recordset
'功能：构建符合门诊处方明细(配药)上传数据结构的记录集对象
'参数：
'  rsVal：HIS门诊处方明细记录集对象
'  lngModule：HIS业务模块号
    
    Dim rsData As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim i As Integer
    Dim strTitle As String, strDetail As String
    Dim lng库房ID As Long
    Dim int单据 As Integer
    Dim strNO As String
    Dim cur应收金额 As Currency, cur实收金额 As Currency
    Dim lngDeviceID As Long
    Dim strTmp As String
    Dim bytType As Byte
    
    If rsVal Is Nothing Then Exit Function
    
    '初始化内存记录集对象
    With rsData
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "DeviceID", adBigInt, , adFldIsNullable
        .Fields.Append "NO", adVarChar, 20, adFldIsNullable
        .Fields.Append "单据", adInteger, , adFldIsNullable
        .Fields.Append "Content", adLongVarChar, 20000, adFldIsNullable
        .Open
    End With
    
    With rsTmp
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "NO", adVarChar, 20, adFldIsNullable
        .Fields.Append "单据", adInteger, , adFldIsNullable
        .Fields.Append "库房ID", adBigInt, , adFldIsNullable
        .Fields.Append "DeviceID", adBigInt, , adFldIsNullable
        .Fields.Append "Type", adInteger, 1, adFldIsNullable
        .Fields.Append "Content", adVarChar, 2000, adFldIsNullable
        .Fields.Append "应收金额", adCurrency, , adFldIsNullable
        .Fields.Append "实收金额", adCurrency, , adFldIsNullable
        .Open
    End With
    
    With rsVal
        If .State <> adStateOpen Then .Open
        i = 1: cur应收金额 = 0: cur实收金额 = 0: strDetail = ""
        
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False

            '符合条件的设备ID
            lngDeviceID = GetDevice(1, !发药药房id, !药品剂型)
            
            If lngDeviceID <= 0 Then GoTo makLoop
            
            bytType = GetDeviceType(lngDeviceID)
            
            '明细信息
            strDetail = ""
            Select Case bytType
            Case enuLinkType.DB
                strDetail = ""
                
            Case enuLinkType.WEBServices
                strDetail = ""
                
            Case enuLinkType.Directory
                strDetail = ""
            End Select
            
            '写入，供rsData记录生成使用
            If strDetail <> "" Then
                rsTmp.AddNew
                rsTmp!NO = !NO
                rsTmp!单据 = !单据
                rsTmp!库房id = !发药药房id
                rsTmp!DeviceID = lngDeviceID
                rsTmp!Type = bytType
                rsTmp!Content = strDetail
                rsTmp!应收金额 = NVL(!应收金额, 0)
                rsTmp!实收金额 = NVL(!实收金额, 0)
                rsTmp.Update
            End If
            
            i = i + 1
            int单据 = !单据: strNO = !NO: lng库房ID = !发药药房id
            
            .MoveNext
            If .EOF Then
                GoTo makCommon1
            ElseIf int单据 <> !单据 And strNO <> !NO And lng库房ID <> !发药药房id Then
makCommon1:
                .MovePrevious
                i = 1
            End If
            
makLoop:
            .MoveNext
        Loop
    End With
    
    '生成最终文本，存入记录集
    With rsTmp
        cur应收金额 = 0
        cur实收金额 = 0
        strDetail = ""
        .Sort = "DeviceID,NO"
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
            strDetail = strDetail & !Content
            lngDeviceID = !DeviceID
            strNO = !NO
            cur应收金额 = cur应收金额 & !应收金额
            cur实收金额 = cur实收金额 & !实收金额
            
            .MoveNext
            If .EOF Then
                GoTo makCommon
            ElseIf lngDeviceID <> !DeviceID And strNO <> !NO Then
makCommon:
                .MovePrevious
                
                '单据文本
                Select Case NVL(!Type, 0)
                Case enuLinkType.DB
                    strTitle = "" & strDetail
                    
                Case enuLinkType.WEBServices
                    strTitle = "" & strDetail
                    
                Case enuLinkType.Directory
                    strTitle = "" & strDetail
                    
                End Select
                
                '加入rsData记录集
                rsData.AddNew
                rsData!DeviceID = lngDeviceID
                rsData!NO = strNO
                rsData!单据 = !单据
                rsData!Content = strTitle
                rsData.Update
                
                strDetail = ""
                cur应收金额 = 0
                cur实收金额 = 0
            End If
            
            .MoveNext
        Loop
        .Close
    End With
    Set rsTmp = Nothing

    Set BuildReceipDetail = rsData

End Function

Private Function BuildReceipList(ByVal rsVal As ADODB.Recordset, ByVal lngModule As Long) As ADODB.Recordset
'功能：构建符合门诊发药上传数据结构的记录集对象
'参数：
'  rsVal：HIS数据集
'  lngModule：HIS业务模块号

    Dim rsData As New ADODB.Recordset
    Dim strBill As String
    Dim lngDeviceID As Long
    Dim arrDeviceID As Variant
    Dim i As Integer
    
    With rsData
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "DeviceID", adBigInt, , adFldIsNullable
        .Fields.Append "NO", adVarChar, 20, adFldIsNullable
        .Fields.Append "Content", adVarChar, 1000, adFldIsNullable
        .Open
    End With
    
    With rsVal
        If .State <> adStateOpen Then .Open
        
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
            
            arrDeviceID = GetDevices(NVL(!发药药房id, 0))
            
            strBill = ""
            
            '相同的发药药房冗余生成数据
            For i = LBound(arrDeviceID) To UBound(arrDeviceID)
                rsData.AddNew
                rsData!DeviceID = arrDeviceID(i)
                rsData!NO = !NO
                rsData!Content = strBill
                rsData.Update
            Next
            Set arrDeviceID = Nothing
            
            .MoveNext
        Loop
        .Close
    End With
    
    '排序
    rsData.Sort = "NO,DeviceID"
    
    Set BuildReceipList = rsData
    
End Function

Private Function BuildReceipAdviceInf(ByVal rsVal As ADODB.Recordset, ByVal lngModule As Long) As ADODB.Recordset
'功能：构建符合住院医嘱发药上传数据结构的记录集对象
'参数：
'  rsVal：HIS数据集
'  lngModule：HIS业务模块号
    
    Dim rsData As New ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim lngDeviceID As Long
    Dim strTmp As String, strDataA As String, strDataB As String
    Dim intCount As Integer, i As Integer
    Dim strNextTime As String
    
    With rsData
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "DeviceID", adBigInt, , adFldIsNullable
        .Fields.Append "Title", adVarChar, 1000, adFldIsNullable
        .Fields.Append "Detail", adLongVarChar, 10000, adFldIsNullable
        .Fields.Append "领药部门ID", adBigInt, , adFldIsNullable
        .Open
    End With
    
    With rsVal
        If .State <> adStateOpen Then .Open
        
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
            lngDeviceID = GetDevice(2, !发药药房id, !药品剂型)
            
            If lngDeviceID <= 0 Then GoTo makLoop
            
            '频率次数
            '如果是临嘱并且是整包装数量，则不发送到包药机
            If Not (!整包装 = 0 Or !医嘱类型 = "长嘱") Then GoTo makLoop
            
            If Val(NVL(!频率间隔)) = 0 Or NVL(!间隔单位) = "" Or NVL(!执行时间方案) = "" Or !医嘱类型 = "临嘱" Then
                intCount = 1
            Else
                intCount = Val(NVL(!频率次数))
                If intCount = 0 Then
                    strTmp = "Select Zl_Gettransexenumber([1],[2],[3],[4],[5],[6]) 执行次数 From Dual "
                    On Error GoTo errHandle
                    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "取执行次数", _
                                CDate(!开始执行时间), CDate(!首次时间), CDate(!末次时间), Val(!频率间隔), !间隔单位, !执行时间方案)
                    If Not rsTmp.EOF Then
                        intCount = Val(rsTmp.Fields(0).Value)
                    End If
                    rsTmp.Close
                    If intCount = 0 Then
                        intCount = 1
                    End If
                    On Error GoTo 0
                End If
            End If
            
            '明细脚本
            '医嘱药品信息
            strDataA = ""
            
            '拆分单次服用数量
            On Error GoTo errHandle
            strNextTime = Format(!首次时间, "YYYY-MM-DD HH:MM:SS")
            strDataB = ""
            For i = 1 To intCount
                strDataB = strDataB & ""
                
                If i < intCount Then
                    strDataB = strDataB & ""
                    
                    gstrSQL = "Select Zl_Gettransexetime([1],[2],[3],[4],[5]) 下次执行时间 From Dual "
                    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "取下次执行时间", _
                                CDate(!开始执行时间), CDate(strNextTime), Val(!频率间隔), !间隔单位, !执行时间方案)
                    If rsTmp.EOF = False Then
                        strNextTime = Format(rsTmp.Fields(0).Value, "YYYY-MM-DD HH:MM:SS")
                    End If
                    rsTmp.Close
                End If
            Next
            On Error GoTo 0
            
            strDataB = strDataB & ""
            
            '单据脚本
            strDataA = ""
            
            '保存
            rsData.AddNew
            rsData!DeviceID = lngDeviceID
            rsData!Title = strDataA
            rsData!Detail = strDataB
            rsData!领药部门ID = !领药部门ID
            rsData.Update
            
makLoop:
            .MoveNext
        Loop
        .Close
    End With
    
    Set BuildReceipAdviceInf = rsData
    Exit Function
    
    
errHandle:
    gstrMessage = Err.Description
    If gobjComLib.ErrCenter = 1 Then Resume
End Function

Private Function BuildDrugStock(ByVal rsDrugStock As ADODB.Recordset, ByVal objDevice As clsDevice) As ADODB.Recordset
'功能：构建符合上传数据结构的药品库存记录集对象
'参数：
'  rsDrugStock：HIS药品库存记录集对象
'  objDevice：设备对象

    Dim strTmp As String
    Dim rsData As New ADODB.Recordset
    Dim i As Integer

    '初始化内存记录集对象
    With rsData
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "SN", adBigInt, , adFldIsNullable
        .Fields.Append "Drug", adVarChar, 100, adFldIsNullable
        .Fields.Append "Content", adVarChar, 3000, adFldIsNullable
        .Open
    End With

    With rsDrugStock
        If .State <> adStateOpen Then .Open
        i = 1
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
        
            '格式化成要上传的数据格式
            Select Case objDevice.LinkType
            Case enuLinkType.DB
                strTmp = ""
            
            Case enuLinkType.WEBServices
                strTmp = ""
                
            Case enuLinkType.Directory
                strTmp = ""
                
            End Select
            
            '存入内存记录集
            If strTmp <> "" Then
                rsData.AddNew
                rsData!SN = i
                rsData!Drug = !编码 & "；" & !通用名 & "；" & NVL(!规格)
                rsData!Content = strTmp
                rsData.Update
                
                i = i + 1
            End If
            
            .MoveNext
        Loop
        .Close
        
    End With
            
    Set BuildDrugStock = rsData

End Function
