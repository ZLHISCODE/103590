Attribute VB_Name = "mdlPiva"
Option Explicit

Private mrsTrans As ADODB.Recordset             '输液单记录，包含输液单内容（药品）
Private mrsPRI As Recordset                     '输液药品优先级
Private mrsVol As Recordset                     '科室容量设置
Private mrstemp As Recordset                    '临时记录
Private mblnLastBatch As Boolean                '是否保持上次批次
Private Sub Piva_GetPara()
    '取配置中心一些参数
    If mrsPRI Is Nothing Then
        gstrSQL = "select 科室id,科室名称,配药类型,频次,有效,优先级 from 输液药品优先级 order by 优先级"
        Set mrsPRI = zlDatabase.OpenSQLRecord(gstrSQL, "获取优先级数据")
    End If
    
    If mrsVol Is Nothing Then
        gstrSQL = "select 科室id,科室名称,容量,配药批次 from 科室容量设置"
        Set mrsVol = zlDatabase.OpenSQLRecord(gstrSQL, "获取科室容量数据")
    End If
    
    mblnLastBatch = (Val(zlDatabase.GetPara("保持上次批次", glngSys, 1345, 0)) = 1)
End Sub


Public Function PIVA_AutoSetBatch(ByVal lng库房id As Long, ByVal strSendNO As String) As Boolean
    '自动设置批次
    'lng库房id：配置中心部门id
    'strdSendNO：医嘱发送号
    Dim rsTrans As ADODB.Recordset
    Dim lng配药id As Long
    Dim int配药类型 As Integer
    Dim rstemp As Recordset
    Dim strOld执行时间 As String
    Dim strOld批次 As String
    Dim lngOld病人id As Long
    Dim lngOld配药id As Long
    Dim lng单量 As Long
    Dim lngCount As Long
    Dim lng优先级 As Long
    Dim int序号 As Integer
    Dim i As Integer
    Dim strInput As String
    Dim arrExecute As Variant
    
    On Error GoTo errHandle
    
    Call Piva_GetPara
    Call Piva_IniTransRec
    Call Piva_IniPriRec
    
    Set rsTrans = Piva_GetTrans(lng库房id, strSendNO)
    
    With rsTrans
        Set rstemp = rsTrans
        If .RecordCount > 0 Then
            rsTrans.Sort = "床号,病人id,配药id,执行时间,批次"
        End If
        Do While Not .EOF
            lngCount = lngCount + 1
            
            '根据优先级和容量规则，设置批次和优先级
            If mrsPRI.RecordCount > 0 Or mrsVol.RecordCount > 0 Then
                
                If strOld执行时间 <> IIf(IsNull(!执行时间), "", Format(!执行时间, "YYYY-MM-DD")) Then
                    '处理上一个病人的容量
                    If lngCount > 1 Then
                        If mrsPRI.RecordCount > 0 Then
                            Call Piva_Set优先级(mrstemp, mrsTrans, lngOld配药id)
                        End If
                        
                        If mrsVol.RecordCount > 0 Then
                            Call Piva_Set批次(mrsTrans, lngOld病人id, strOld批次, strOld执行时间)
                        End If
                    End If
                    
                    '当前病人的容量
                    Call Piva_IniPriRec
                    strOld执行时间 = IIf(IsNull(!执行时间), "", Format(!执行时间, "YYYY-MM-DD"))
                    strOld批次 = IIf(IsNull(!配药批次), "", !配药批次 & "#")
                    lngOld病人id = IIf(IsNull(!病人ID), 0, !病人ID)
                    lngOld配药id = !配药id
                Else
                    If strOld批次 <> IIf(IsNull(!配药批次), "", !配药批次 & "#") Then
                        '处理上一个病人的容量
                        If mrsPRI.RecordCount > 0 Then
                            Call Piva_Set优先级(mrstemp, mrsTrans, lngOld配药id)
                        End If
                        
                        If mrsVol.RecordCount > 0 Then
                            Call Piva_Set批次(mrsTrans, lngOld病人id, strOld批次, strOld执行时间)
                        End If
                        '当前病人的容量
                        Call Piva_IniPriRec
                        strOld批次 = IIf(IsNull(!配药批次), "", !配药批次 & "#")
                        lngOld病人id = IIf(IsNull(!病人ID), 0, !病人ID)
                        lngOld配药id = !配药id
                        
                    Else
                        If lngOld病人id <> IIf(IsNull(!病人ID), 0, !病人ID) Then
                            '处理上一个病人的容量
                            If mrsPRI.RecordCount > 0 Then
                                Call Piva_Set优先级(mrstemp, mrsTrans, lngOld配药id)
                            End If
                            
                            If mrsVol.RecordCount > 0 Then
                                Call Piva_Set批次(mrsTrans, lngOld病人id, strOld批次, strOld执行时间)
                            End If
                            '当前病人的容量
                            Call Piva_IniPriRec
                            lngOld病人id = IIf(IsNull(!病人ID), 0, !病人ID)
                            lngOld配药id = !配药id
                        Else
                            If lngOld配药id <> !配药id Then
                                If Not mrsPRI.RecordCount Then
                                    Call Piva_Set优先级(mrstemp, mrsTrans, lngOld配药id)
                                End If
                                lngOld配药id = !配药id
                            End If
                        End If
                    End If
                End If
                
                '保存数据集
                mrstemp.AddNew
                mrstemp!部门ID = !病人科室id
                mrstemp!配药id = !配药id
                mrstemp!配药类型 = IIf(IsNull(!配药类型), "", !配药类型)
                mrstemp!频次 = IIf(IsNull(!执行频次), "", !执行频次)
                mrstemp.Update
            End If
             
            mrsTrans.AddNew
            mrsTrans!配药id = !配药id
            mrsTrans!部门ID = !部门ID
            mrsTrans!序号 = !序号
            mrsTrans!姓名 = IIf(IsNull(!姓名), "", !姓名)
            mrsTrans!性别 = IIf(IsNull(!性别), "", !性别)
            mrsTrans!年龄 = IIf(IsNull(!年龄), "", !年龄)
            mrsTrans!住院号 = IIf(IsNull(!住院号), "", !住院号)
            mrsTrans!床号 = IIf(IsNull(!床号), "", !床号)
            mrsTrans!病人病区 = !病人病区
            mrsTrans!病人科室 = !病人科室
            mrsTrans!执行时间 = IIf(IsNull(!执行时间), "", Format(!执行时间, "YYYY-MM-DD HH:MM:SS"))
            mrsTrans!病人ID = IIf(IsNull(!病人ID), 0, !病人ID)
            mrsTrans!主页id = IIf(IsNull(!主页id), 0, !主页id)
            mrsTrans!病人科室id = IIf(IsNull(!病人科室id), 0, !病人科室id)
            mrsTrans!打包时间 = IIf(IsNull(!打包时间), "", !打包时间)
            
            mrsTrans!配药批次 = IIf(IsNull(!配药批次), "", !配药批次 & "#")
            mrsTrans!新配药批次 = IIf(IsNull(!配药批次), "", !配药批次 & "#")
            mrsTrans!瓶签号 = IIf(IsNull(!瓶签号), "", !瓶签号)
            mrsTrans!打印标志 = IIf(IIf(IsNull(!打印标志), 0, !打印标志) = 0, 0, 1)
            mrsTrans!是否打包 = IIf(IsNull(!是否打包), 0, !是否打包)
            mrsTrans!核查人 = IIf(IsNull(!操作人员), "", !操作人员)
            mrsTrans!核查时间 = IIf(IsNull(!操作时间), "", Format(!操作时间, "YYYY-MM-DD HH:MM:SS"))
            mrsTrans!摆药人 = IIf(IsNull(!操作人员), "", !操作人员)
            mrsTrans!摆药时间 = IIf(IsNull(!操作时间), "", Format(!操作时间, "YYYY-MM-DD HH:MM:SS"))
            mrsTrans!摆药单号 = IIf(IsNull(!摆药单号), "", !摆药单号)
            mrsTrans!配药人 = IIf(IsNull(!操作人员), "", !操作人员)
            mrsTrans!配药时间 = IIf(IsNull(!操作时间), "", Format(!操作时间, "YYYY-MM-DD HH:MM:SS"))
            mrsTrans!发送人 = IIf(IsNull(!操作人员), "", !操作人员)
            mrsTrans!发送时间 = IIf(IsNull(!操作时间), "", Format(!操作时间, "YYYY-MM-DD HH:MM:SS"))
            mrsTrans!销帐申请人 = IIf(IsNull(!操作人员), "", !操作人员)
            mrsTrans!销帐申请时间 = IIf(IsNull(!操作时间), "", Format(!操作时间, "YYYY-MM-DD HH:MM:SS"))
            mrsTrans!销帐审核人 = IIf(IsNull(!操作人员), "", !操作人员)
            mrsTrans!销帐审核时间 = IIf(IsNull(!操作时间), "", Format(!操作时间, "YYYY-MM-DD HH:MM:SS"))
            mrsTrans!抗菌药物 = 1
            mrsTrans!药师审核时间 = IIf(IsNull(!药师审核时间), 0, !药师审核时间)
            mrsTrans!是否调整批次 = IIf(IsNull(!是否调整批次), 0, !是否调整批次)
            mrsTrans!是否锁定 = IIf(IsNull(!是否锁定), 0, !是否锁定)
            mrsTrans!手工调整批次 = IIf(IsNull(!手工调整批次), 0, !手工调整批次)
            mrsTrans!拒收原因 = IIf(IsNull(!拒收原因), "", !拒收原因)
            
            mrsTrans!收发Id = !收发Id
            mrsTrans!单据 = !单据
            mrsTrans!NO = !NO
            mrsTrans!药品名称 = "[" & !药品编码 & "]" & !通用名
            mrsTrans!通用名 = !通用名
            mrsTrans!商品名 = IIf(IsNull(!商品名), "", !商品名)
            mrsTrans!英文名 = IIf(IsNull(!英文名), "", !英文名)
            mrsTrans!规格 = IIf(IsNull(!规格), "", !规格)
            mrsTrans!产地 = IIf(IsNull(!产地), "", !产地)
            mrsTrans!批号 = IIf(IsNull(!批号), "", !批号)
            mrsTrans!单量 = IIf(IsNull(!单量), 0, !单量)
            mrsTrans!剂量单位 = !剂量单位
            mrsTrans!频次 = IIf(IsNull(!频次), "", !频次)
            mrsTrans!数量 = IIf(IsNull(!数量), 0, !数量)
            mrsTrans!单位 = !单位
            mrsTrans!批次 = !批次
            mrsTrans!用法 = IIf(IsNull(!用法), "", !用法)
            mrsTrans!药品ID = IIf(IsNull(!药品ID), 0, !药品ID)
            mrsTrans!药名ID = !药名ID
            mrsTrans!费用序号 = !费用序号
            mrsTrans!费用ID = !费用ID
            mrsTrans!配药类型 = !配药类型
            
            mrsTrans!发药数量 = IIf(IsNull(!发药数量), 0, !发药数量)
            mrsTrans!库存数量 = IIf(IsNull(!库存数量), 0, !库存数量)
            mrsTrans!实际数量 = IIf(IsNull(!实际数量), 0, !库存数量)
            
            mrsTrans!医嘱id = !医嘱id
            mrsTrans!发送号 = !发送号
            mrsTrans!医嘱发送时间 = IIf(IsNull(!医嘱发送时间), "", Format(!医嘱发送时间, "YYYY-MM-DD HH:MM:SS"))
            mrsTrans!审查结果 = IIf(IsNull(!审查结果), 0, !审查结果)
            mrsTrans!作废类型 = IIf(IsNull(!作废类型), "", !作废类型)
            mrsTrans!险类 = !险类
            mrsTrans!颜色 = !颜色
            mrsTrans!执行标志 = 0
            mrsTrans!溶媒 = IIf(IsNull(!溶媒), 0, !溶媒)
            
            If !配药id <> lng配药id Then
                int序号 = int序号 + 1
            End If
            mrsTrans!组号 = int序号
            mrsTrans.Update
            

            If lngCount = .RecordCount Then
                If mrsPRI.RecordCount > 0 Then
                    Call Piva_Set优先级(mrstemp, mrsTrans, lngOld配药id)
                End If
                
                If mrsVol.RecordCount > 0 Then
                    Call Piva_Set批次(mrsTrans, lngOld病人id, strOld批次, strOld执行时间)
                End If
            End If

            If !配药id = lng配药id Then
                If Val(!配药类型) > 0 Then
                    int配药类型 = 1
                ElseIf int配药类型 = 0 And Val(!配药类型) = 0 Then
                    int配药类型 = 0
                End If
            Else
                int配药类型 = Val(!配药类型)
            End If
            
            mrsTrans.Filter = "配药id=" & !配药id
            Do While Not mrsTrans.EOF
                mrsTrans.Update "抗菌药物", int配药类型
                mrsTrans.MoveNext
            Loop
            
            mrsTrans.Filter = ""
            lng配药id = !配药id
            
            .MoveNext
        Loop
    End With
    
    lng配药id = 0
    
    '保持上次批次
    If mblnLastBatch = True Then
        Call Piva_SetLastBatch(mrsTrans)
    End If
    
    With mrsTrans
        .Filter = ""
        .Sort = "配药ID"
        Do While Not .EOF
            If !新配药批次 <> !配药批次 And lng配药id <> Val(!配药id) Then
                lng配药id = Val(!配药id)
                
                If IIf(IsNull(!新配药批次), "", !新配药批次) = "" Then
                    strInput = IIf(strInput = "", "", strInput & "|") & !配药id & ",:" & IIf(IsNull(!优先级), 0, !优先级)
                Else
                    strInput = IIf(strInput = "", "", strInput & "|") & !配药id & "," & Mid(!新配药批次, 1, IIf(Len(!新配药批次) = 0, 0, Len(!新配药批次) - 1)) & ":" & IIf(IsNull(!优先级), 0, !优先级)
                End If
            End If
            .MoveNext
        Loop
    End With
    
    If strInput <> "" Then
        arrExecute = Piva_GetArrayByStr(strInput, 3900, "|")
        For i = 0 To UBound(arrExecute)
            gstrSQL = "Zl_输液配药记录_分批("
            '配药ID,批次
            gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
            gstrSQL = gstrSQL & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "Piva_AutoSetBatch")
        Next
    End If
    
    PIVA_AutoSetBatch = True
    
    Exit Function
errHandle:
    PIVA_AutoSetBatch = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Piva_GetArrayByStr(ByVal strInput As String, ByVal lngLength As Long, ByVal strSplitChar As String) As Variant
    '根据传入的字符串进行分解，大于指定字符长度就需要进行分解，结果保存到数组中
    '入参：strInput-输入的字符串；strSplitChar-字符串中内容的分隔符
    '返回：数组，其中数组成员的字符长度不超过指定长度
    Dim strArray As Variant
    Dim ArrTmp As Variant
    Dim strTmp As String
    Dim lngCount As Long
    Dim i As Long
    
    strArray = Array()
   
    '大于指定字符时就需要分解
    If Len(strInput) > lngLength Then
        If strSplitChar = "" Then
            '无分隔符时
            strTmp = strInput
            Do While Len(strTmp) > lngLength
                ReDim Preserve strArray(UBound(strArray) + 1)
                strArray(UBound(strArray)) = Mid(strTmp, 1, lngLength)
                strTmp = Mid(strTmp, lngLength + 1)
            Loop
            
            If strTmp <> "" Then
                ReDim Preserve strArray(UBound(strArray) + 1)
                strArray(UBound(strArray)) = strTmp
            End If
        Else
            '有分隔符时
            ArrTmp = Split(strInput & strSplitChar, strSplitChar)
            lngCount = UBound(ArrTmp)
        
            For i = 0 To lngCount
                If ArrTmp(i) <> "" Then
                    '有分隔符的需要保持分隔符之间字符的完整性，不能把分隔符之间的字符拆开
                    If Len(IIf(strTmp = "", "", strTmp & strSplitChar) & ArrTmp(i)) > lngLength Then
                        ReDim Preserve strArray(UBound(strArray) + 1)
                        strArray(UBound(strArray)) = strTmp
                        strTmp = ArrTmp(i)
                    Else
                        strTmp = IIf(strTmp = "", "", strTmp & strSplitChar) & ArrTmp(i)
                    End If
                End If
                       
                If i = lngCount Then
                    ReDim Preserve strArray(UBound(strArray) + 1)
                    strArray(UBound(strArray)) = strTmp
                End If
            Next
        End If
    Else
        ReDim Preserve strArray(UBound(strArray) + 1)
        strArray(UBound(strArray)) = strInput
    End If
    
    Piva_GetArrayByStr = strArray
End Function
Private Sub Piva_IniTransRec()
    '输液单记录集
    Set mrsTrans = New ADODB.Recordset
    With mrsTrans
        If .State = 1 Then .Close
        
        '该记录对应的输液配药记录信息
        .Fields.Append "组号", adDouble, 18, adFldIsNullable
        .Fields.Append "配药id", adDouble, 18, adFldIsNullable
        .Fields.Append "部门id", adDouble, 18, adFldIsNullable
        .Fields.Append "序号", adDouble, 3, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "性别", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "年龄", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "住院号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "床号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "病人病区", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "病人科室", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "执行时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "病人id", adDouble, 18, adFldIsNullable
        .Fields.Append "主页id", adDouble, 18, adFldIsNullable
        .Fields.Append "优先级", adDouble, 18, adFldIsNullable
        .Fields.Append "病人科室id", adDouble, 18, adFldIsNullable
        .Fields.Append "打包时间", adLongVarChar, 20, adFldIsNullable
        
        '输液配药记录业务操作信息
        .Fields.Append "配药批次", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "瓶签号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "打印标志", adDouble, 1, adFldIsNullable
        .Fields.Append "是否打包", adDouble, 1, adFldIsNullable
        .Fields.Append "核查人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "核查时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "摆药人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "摆药时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "摆药单号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "配药人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "配药时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "发送人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "发送时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "销帐申请人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "销帐申请时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "销帐审核人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "销帐审核时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "抗菌药物", adLongVarChar, 1, adFldIsNullable
        .Fields.Append "药师审核时间", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "是否调整批次", adDouble, 1, adFldIsNullable
        .Fields.Append "是否锁定", adDouble, 1, adFldIsNullable
        .Fields.Append "新配药批次", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "手工调整批次", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "拒收原因", adLongVarChar, 200, adFldIsNullable
        
        '输液配药记录对应的药品信息
        .Fields.Append "收发id", adDouble, 18, adFldIsNullable
        .Fields.Append "单据", adDouble, 2, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "药品名称", adLongVarChar, 50, adFldIsNullable   '编码+通用名/商品名
        .Fields.Append "通用名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "商品名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "英文名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "单量", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "剂量单位", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "频次", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "数量", adDouble, 18, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .Fields.Append "用法", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "药名id", adDouble, 18, adFldIsNullable
        .Fields.Append "费用序号", adDouble, 3, adFldIsNullable
        .Fields.Append "费用id", adDouble, 18, adFldIsNullable
        .Fields.Append "配药类型", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "溶媒", adDouble, 1, adFldIsNullable
        
        .Fields.Append "发药数量", adDouble, 18, adFldIsNullable
        .Fields.Append "库存数量", adDouble, 18, adFldIsNullable
        .Fields.Append "实际数量", adDouble, 18, adFldIsNullable
        
        .Fields.Append "审查结果", adDouble, 1, adFldIsNullable
        .Fields.Append "医嘱发送时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "医嘱id", adDouble, 18, adFldIsNullable
        .Fields.Append "发送号", adDouble, 18, adFldIsNullable
        
        .Fields.Append "执行标志", adDouble, 1, adFldIsNullable
        .Fields.Append "作废类型", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "险类", adDouble, 5, adFldIsNullable
        .Fields.Append "颜色", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub Piva_IniPriRec()
    Set mrstemp = New ADODB.Recordset
    With mrstemp
        If .State = 1 Then .Close
        .Fields.Append "配药id", adDouble, 18, adFldIsNullable
        .Fields.Append "部门id", adDouble, 18, adFldIsNullable
        .Fields.Append "配药类型", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "频次", adLongVarChar, 20, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub Piva_Set优先级(ByVal rstemp As Recordset, ByRef rsTrans As Recordset, ByVal lng配药id As Long)
    Dim lng优先级A As Long
    Dim lng优先级B As Long
    Dim lng优先级C As Long
    Dim lng优先级D As Long
    Dim bln其他类型 As Boolean
    Dim bln其他频次 As Boolean
    
    If rstemp.EOF Or mrsPRI.EOF Then Exit Sub
    mrsPRI.MoveFirst
    mrsPRI.Sort = "优先级"
    
    mrsPRI.Filter = "科室id='" & rstemp!部门ID & "'"

    rsTrans.Filter = "配药id=" & lng配药id
    rsTrans.Sort = ""
    
    If mrsPRI.EOF Then mrsPRI.Filter = "科室id='0'"
    Do While Not mrsPRI.EOF
        rsTrans.MoveFirst
        Do While Not rsTrans.EOF
            If mrsPRI!配药类型 = rsTrans!配药类型 Then
                bln其他类型 = True
                If Mid(mrsPRI!频次, 1, IIf(InStr(1, mrsPRI!频次, "(") = 0, 1, InStr(1, mrsPRI!频次, "(") - 1)) = rsTrans!频次 Then
                    bln其他频次 = True
                    lng优先级A = mrsPRI!优先级
                ElseIf mrsPRI!频次 = "其他频次" And Not bln其他频次 Then
                    lng优先级B = mrsPRI!优先级
                ElseIf mrsPRI!频次 = "所有频次" And Not bln其他频次 Then
                    lng优先级B = mrsPRI!优先级
                End If
            ElseIf mrsPRI!配药类型 = "其他类型" And Not bln其他类型 Then
                If Mid(mrsPRI!频次, 1, IIf(InStr(1, mrsPRI!频次, "(") = 0, 1, InStr(1, mrsPRI!频次, "(") - 1)) = rsTrans!频次 Then
                    bln其他频次 = True
                    lng优先级C = mrsPRI!优先级
                ElseIf mrsPRI!频次 = "其他频次" And Not bln其他频次 Then
                    lng优先级D = mrsPRI!优先级
                ElseIf mrsPRI!频次 = "所有频次" And Not bln其他频次 Then
                    lng优先级D = mrsPRI!优先级
                End If
            End If
            rsTrans.MoveNext
        Loop
        mrsPRI.MoveNext
    Loop
    
    rsTrans.MoveFirst
    Do While Not rsTrans.EOF
        rsTrans!优先级 = IIf(bln其他类型, IIf(bln其他频次, lng优先级A, lng优先级B), IIf(bln其他频次, lng优先级C, lng优先级D))
        rsTrans.Update
        rsTrans.MoveNext
    Loop
    rsTrans.Filter = ""
End Sub

Private Function Piva_GetTrans(ByVal lngCenterID As Long, ByVal strSendNO As String) As ADODB.Recordset
    '取输液配药记录
    'lngCenterID：输液配置中心ID
    'str病区ID：病区ID串
    'dateExeStart、dateExeEnd：输液配药单据的执行时间范围
    On Error GoTo errHandle
    
    gstrSQL = "Select Distinct A.ID As 配药ID, A.部门id, A.序号, A.配药批次, S.颜色,A.姓名, A.性别, A.年龄, A.住院号, A.床号,M.药师审核时间, A.病人病区id, A.病人科室id, A.执行时间, A.瓶签号,A.打包时间,M.执行频次,A.是否调整批次,A.是否锁定,A.手工调整批次,'' 拒收原因," & _
        "  A.操作人员,A.操作时间,Nvl(A.打印标志,0) As 打印标志, A.是否打包, B.名称 As 病人病区, C.名称 As 病人科室, D.收发id, E.单据, E.NO, F.编码 As 药品编码, " & _
        " F.名称 As 通用名, H.名称 As 商品名, I.名称 As 英文名, F.规格, E.产地, E.批号, E.单量, J.计算单位 As 剂量单位,J.id 药名id, E.频次, '' As 作废类型, " & _
        " Case Nvl(E.审核人, '未审核') When '未审核' Then E.实际数量 * Nvl(E.付数, 1) / G.住院包装 Else 0 End As 发药数量,M.病人id,M.主页id,T.溶媒,A.医嘱id,A.发送号, " & _
        " (D.数量 / G.住院包装)  As 数量,D.数量 As 实际数量, G.住院单位 As 单位,Nvl(E.批次,0) As 批次, Nvl(L.实际数量, 0)/ G.住院包装 As 库存数量, Nvl(M.审查结果,-1) 审查结果, E.用法, E.药品id, n.序号 As 费用序号,E.费用id, o.险类, A.摆药单号,r.发送时间 As 医嘱发送时间,nvl(T.抗生素,'0') 配药类型 " & _
        " From  输液配药记录 A, 部门表 B, 部门表 C, 输液配药内容 D, 药品收发记录 E, 收费项目目录 F, 药品规格 G,输液药品属性 X,  收费项目别名 H, 诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M, 住院费用记录 N, 病案主页 O ,配药工作批次 S,药品特性 T " & _
        ",(Select 库房id, 药品id, Nvl(批次, 0) As 批次, Nvl(实际数量, 0) As 实际数量 " & _
        " From 药品库存 Where 性质 = 1 And 库房id = [1]) L, 药品收发记录 P, 病人医嘱发送 R " & _
        " Where A.病人病区id = B.ID And A.病人科室id = C.ID And A.ID = D.记录id And D.收发id = E.ID And E.药品id = F.ID And F.ID = G.药品id And G.药品id=X.药品id(+) And E.费用id = N.ID And N.医嘱序号 = M.ID And " & _
        " G.药品id = H.收费细目id(+) And H.性质(+) = 3 And G.药名id = I.诊疗项目id(+) And I.性质(+) = 2 And G.药名id = J.ID And T.药名id=J.ID And A.配药批次=S.批次(+) And E.库房id = L.库房id(+) And E.药品id = L.药品id(+) And Nvl(E.批次, 0) = L.批次(+) " & _
        " And n.病人id = o.病人id(+) And n.主页id = o.主页id(+) And A.部门id = [1] And a.医嘱id = r.医嘱id And a.发送号 = r.发送号 " & _
        " And e.单据 = p.单据 And e.No = p.No And e.库房id + 0 = p.库房id And e.药品id + 0 = p.药品id+0 And e.序号 = p.序号 And (p.记录状态 = 1 Or Mod(p.记录状态, 3) = 0) " & _
        " And A.操作状态=1 And R.发送号 = [2] "
     
    Set Piva_GetTrans = zlDatabase.OpenSQLRecord(gstrSQL, "读取输液配药记录", lngCenterID, strSendNO)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub Piva_Set批次(ByRef rstemp As Recordset, ByVal lng病人ID As Long, ByVal str批次 As String, ByVal str执行时间 As String)
    Dim lng容量 As Long
    Dim lng总量 As Long
    Dim lngRow As Long
    Dim bln批次 As Boolean
    Dim lng配药id As Long
    Dim rs病人 As Recordset
    Dim str配药批次 As String
    Dim strCon As String
    Dim lngOld配药id As Long
    Dim blnLoop As Boolean
    Dim lng单组 As Long
    Dim strOld批次 As String
    Dim strC批次 As String
    
    If rstemp.EOF Then Exit Sub
    
    rstemp.MoveFirst
    rstemp.Sort = "新配药批次,优先级,执行时间,配药id"
    
    rstemp.Filter = "病人id=" & lng病人ID & " "
    rstemp.MoveFirst
    Do While Not rstemp.EOF
        blnLoop = False
        If Format(rstemp!执行时间, "YYYY-MM-DD") = str执行时间 Then
            If rstemp!新配药批次 = strOld批次 Then
                If rstemp!溶媒 = 1 Then lng总量 = lng总量 + rstemp!单量
            Else
                strOld批次 = rstemp!新配药批次
                lng容量 = 0
                mrsVol.MoveFirst
                Do While Not mrsVol.EOF
                    If (mrsVol!科室ID = "0" Or Val(mrsVol!科室ID) = rstemp!病人科室id Or mrsVol!科室ID = "00") And (mrsVol!配药批次 = rstemp!新配药批次 Or mrsVol!配药批次 = "") Then
                        If Val(mrsVol!容量) > lng容量 Then
                            lng容量 = Val(mrsVol!容量)
                        End If
                    End If
                    mrsVol.MoveNext
                Loop
                
                lng总量 = 0
                lng配药id = 0
                If rstemp!溶媒 = 1 Then lng总量 = rstemp!单量
            End If
            
            
            If lng配药id <> rstemp!配药id Then
                lng单组 = 0
                strCon = strCon & " And 配药id<>" & rstemp!配药id
                If rstemp!溶媒 = 1 Then lng单组 = rstemp!单量
                lng配药id = rstemp!配药id
                lngRow = lngRow + 1
            Else
                If rstemp!溶媒 = 1 Then lng单组 = lng单组 + rstemp!单量
            End If
        
            If lngRow > 1 And lng容量 > 0 Then
                If lng总量 - lng单组 > lng容量 And lng单组 <> 0 Then
                    lng总量 = lng总量 - lng单组
                    rstemp.Filter = "配药id=" & lng配药id & ""
                    rstemp.Sort = ""
                     
                    rstemp.MoveFirst
                    Do While Not rstemp.EOF
                        rstemp!新配药批次 = Val(Mid(rstemp!新配药批次, 1, Len(rstemp!新配药批次) - 1)) + 1 & "#"
                        rstemp.Update
                        rstemp.MoveNext
                    Loop
                    
                    strC批次 = strOld批次
                    If strCon <> "" Then strCon = Mid(strCon, 1, Len(strCon) - Len(" And 配药id<>" & lng配药id))
                    
                    rstemp.Filter = "病人id=" & lng病人ID & strCon
                    rstemp.Sort = "新配药批次,优先级,执行时间,配药id"
                    If rstemp.RecordCount <> 0 Then rstemp.MoveFirst
                    blnLoop = True
                End If
            End If
        Else
            strCon = strCon & " And 配药id<>" & rstemp!配药id
        End If
        If Not rstemp.EOF And Not blnLoop Then rstemp.MoveNext
    Loop
    rstemp.Filter = ""
End Sub



Private Sub Piva_SetLastBatch(ByRef rsTrans As ADODB.Recordset)
    '保持上次批次，当医嘱变化(新开,新停)时不保持
    Dim lng医嘱id As Long
    Dim lng发送号 As Long
    Dim str执行时间 As String
    Dim rsData As ADODB.Recordset
    Dim strOldFilter, strOldSort As String
    Dim str病人ids As String
    Dim lng病人ID As Long
    Dim str保持批次病人ids As String
    Dim str本次医嘱ids As String
    Dim str上次医嘱ids As String
    Dim intCount As Integer
    Dim i As Integer
    
    '记录初始的过滤及排序内容
    strOldFilter = rsTrans.Filter
    strOldSort = rsTrans.Sort
    
    '统计病人
    With rsTrans
        .Filter = ""
        If .RecordCount = 0 Then Exit Sub
        .Sort = "病人id"
        
        Do While Not .EOF
            If lng病人ID <> !病人ID Then
                lng病人ID = !病人ID
            
                gstrSQL = "Select a.医嘱id, a.发送号 From 输液配药记录 A, 病人医嘱记录 B,诊疗项目目录 C Where a.医嘱id = b.Id and B.诊疗项目id=C.id and c.操作类型=2 and b.诊疗类别='E' And b.病人id = [1] and b.主页id =[2]"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "SetLastBatch", lng病人ID, !主页id)
                
                str本次医嘱ids = ""
                str上次医嘱ids = ""
                lng发送号 = 0
                intCount = 0
                With rsData
                    .Sort = "发送号 Desc,医嘱id"
                    
                    Do While Not .EOF
                        If lng发送号 <> !发送号 Then
                            intCount = intCount + 1
                            lng发送号 = !发送号
                        End If
                        
                        If intCount = 1 Then
                            '取本次发送的医嘱ID
                            If InStr(1, str本次医嘱ids, !医嘱id) = 0 Then
                                str本次医嘱ids = IIf(str本次医嘱ids = "", "", str本次医嘱ids & ",") & !医嘱id
                            End If
                        ElseIf intCount = 2 Then
                            '取上次发送的医嘱ID
                            If InStr(1, str上次医嘱ids, !医嘱id) = 0 Then
                                str上次医嘱ids = IIf(str上次医嘱ids = "", "", str上次医嘱ids & ",") & !医嘱id
                            End If
                        Else
                            Exit Do
                        End If
                        
                        rsData.MoveNext
                    Loop
                End With
            
                '两次发送医嘱的医嘱ID一样，表示没有变化
                If str本次医嘱ids = str上次医嘱ids Then
                    str保持批次病人ids = IIf(str保持批次病人ids = "", "", str保持批次病人ids & ",") & lng病人ID
                End If
            End If
        .MoveNext
        Loop
    End With
    
    If str保持批次病人ids = "" Then Exit Sub
    
    '保持上次批次设置
    For i = 0 To UBound(Split(str保持批次病人ids, ","))
         With rsTrans
            .Filter = "病人id=" & Split(str保持批次病人ids, ",")(i)
            .Sort = "配药ID"
            
            Do While Not .EOF
                lng医嘱id = !医嘱id
                lng发送号 = !发送号
                str执行时间 = IIf(IsNull(!执行时间), "", Format(!执行时间, "YYYY-MM-DD HH:MM:SS"))
    
                gstrSQL = " Select Distinct 配药批次 " & _
                    " From 输液配药记录 A " & _
                    " Where 医嘱id = [1] And 发送号 = (Select Distinct Max(发送号) From 输液配药记录 Where 医嘱id = [1] And 发送号 <> [2]) And " & _
                    " To_Char(a.执行时间, 'hh24:mi:ss') = To_Char([3], 'hh24:mi:ss') "
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "SetLastBatch", lng医嘱id, lng发送号, CDate(str执行时间))
                
                If rsData.RecordCount > 0 Then
                    !新配药批次 = rsData!配药批次 & "#"
                    .Update
                End If
                
                .MoveNext
            Loop
        End With
    Next
    
    '恢复记录集状态
    rsTrans.Filter = IIf(strOldFilter = "0", 0, strOldFilter)
    rsTrans.Sort = strOldSort
End Sub



