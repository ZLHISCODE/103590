Attribute VB_Name = "mdlParameter"
Option Explicit
Public Const gstrParSplit1 As String = "^"  '批量保存参数的模块、参数号、参数值分隔符
Public Const gstrParSplit2 As String = "#"  '批量保存参数组分隔符

Public Enum Enum_Module
    P病人入院管理 = 1131
    p病人入出管理 = 1132
    p病人信息管理 = 1101
    p门诊病历管理 = 1250
    p住院病历管理 = 1251
    p门诊医嘱下达 = 1252
    p住院医嘱下达 = 1253
    p住院医嘱发送 = 1254
    p护理记录管理 = 1255
    p临床路径应用 = 1256
    p门诊路径应用 = 1248
    p临床路径管理 = 1078
    p辅诊记录管理 = 1256
    p医嘱附费管理 = 1257
    p诊疗报告管理 = 1258
    p门诊医生站 = 1260
    p住院医生站 = 1261
    p住院护士站 = 1262
    p医技工作站 = 1263
    p门诊输液管理 = 1264
    p新版住院护士站 = 1265
    p疾病诊断参考 = 1270
    p药品诊疗参考 = 1271
    p病人病历检索 = 1273
    p观片工具管理 = 1289
    
    '药品业务
    p药品目录管理 = 1023
    p药品外购管理 = 1300
    p药品自制入库 = 1301
    p药品移库管理 = 1304
    p药品领用管理 = 1305
    p药品盘点管理 = 1307
    p药品调价管理 = 1333
    p药品质量管理 = 1331
    p药品处方发药 = 1341
    p药品部门发药 = 1342
    p药品申领管理 = 1343
    p大处方审查 = 1347
    p输液配置中心 = 1345

    '病历业务
    p病历内部工具 = 1070
    p电子病案审查 = 1560
    p电子病案借阅 = 1561
    p电子病案评分 = 1562
    
    '费用业务
    p费用虚拟模块 = 9000
    p预交款管理 = 1103
    p医疗卡管理 = 1107
    p挂号安排 = 1110
    p挂号管理 = 1111
    p分诊管理 = 1113
    p临床出诊安排 = 1114
    p门诊划价管理 = 1120
    p门诊收费管理 = 1121
    p门诊记帐管理 = 1122
    p门诊补结算 = 1124
    p住院记帐管理 = 1133
    p科室分散记帐 = 1134
    p医技科室记帐 = 1135
    p住院记帐操作 = 1150
    p病人结帐管理 = 1137
    p执行登记管理 = 1142
    p费用审核管理 = 1143
    p一卡通消费操作 = 1151
    p收费财务监控 = 1500
    p票据使用监控 = 1501
    p人员借款管理 = 1502
    p消费卡管理 = 1503
    p票据入库管理 = 1504
    p收费轧帐管理 = 1506
    
    '处方审查
    p门诊处方审查 = 1351
    p住院药嘱审查 = 1352
    p处方审查项目 = 1353
    p处方审查条件 = 1354
    p处方审查统计 = 1355

    'PACS
    p影像观片设置 = 1288
    p影像医技设置 = 1290
    p影像采集设置 = 1291
    p影像病理设置 = 1294
    p病理归档设置 = 1295
    p病理借还设置 = 1296
End Enum

Public Enum ParaErrType
    PET_正常 = 0
    PET_参数丢失 = 1 '该参数不存在
    PET_不能设置 = 2 '该参数是私有或本机参数无法在此处进行参数设置
    PET_值超限 = 3 '该参数值超出可控件容许范围
End Enum

Public Sub InitSCBItem(ByRef scb As ShortcutBar, ByVal strItems As String, ByRef lngTPLhwnd As Long, Optional ByVal lngSelectedItem As Long = 1)
'功能：初始化一个快捷面板分类列表
'参数：
'      strItems         - 多个分类列表名称，以逗号分隔,例：基础数据初始,流程与规则,接口配置
'      lngTPLhwnd       - 分类列表上绑定的TaskPanel所在的容器句柄（窗体或Picture）
'      lngSelectedItem  - 缺省选中项的序号,从1开始

    Dim scbItem As ShortcutBarItem
    Dim i As Long
    Dim arrItem As Variant
    
    arrItem = Split(strItems, ",")
    For i = 0 To UBound(arrItem)
        Set scbItem = scb.AddItem(i + 1, arrItem(i), lngTPLhwnd)    '图标序号比指定的小1，所以要加1
        If i + 1 = lngSelectedItem Then Set scb.Selected = scbItem
    Next
    
    scb.ExpandedLinesCount = scb.ItemCount
End Sub


Public Sub InitTPLItem(ByRef scc As ShortcutCaption, ByRef tplFunc As TaskPanel, _
        ByVal strCategory As String, ByVal strItems As String, Optional ByVal lngSelectedItem As Long = 1)
'功能：初始或重新加载一个任务面板列表（仅一个分组）
'参数：
'      strCategory      - 显示在ShotcutCaption上的当前分类名称
'      strItems         - 多个二级分类的名称，以分号分隔,以逗号分隔图标ID、容器数组及二级分类名称,例：401,1,门诊划价管理;412,2,病人收费管理;......
'      lngSelectedItem  - 缺省选中项的序号,从1开始

    Dim tplGroup As TaskPanelGroup
    Dim tplItem As TaskPanelGroupItem
    Dim arrItem As Variant
    Dim i As Long
    Dim lngImg As Long, lngId As Long
    Dim strItem As String
    Dim lngUbound As Long
    
    '增加一个隐藏分组
    scc.Caption = strCategory
    If tplFunc.Groups.Count = 0 Then
        Set tplGroup = tplFunc.Groups.Add(1, "分组")
        tplGroup.CaptionVisible = False
        tplGroup.Expanded = True
        
        tplFunc.SetMargins 1, 2, 0, 2, 2
        tplFunc.SetIconSize 24, 24
        tplFunc.SelectItemOnFocus = True
    Else
        Set tplGroup = tplFunc.Groups(1)    'index是从1开始的
        tplGroup.Items.Clear
    End If
    
    arrItem = Split(strItems, ";")
    lngUbound = UBound(arrItem)
    For i = 0 To lngUbound
        lngImg = Split(arrItem(i), ",")(0) + 1  '图标序号比指定的小1，所以要加1
        lngId = Split(arrItem(i), ",")(1)       'ID（作为参数控件容器的Picture数组编号）
        strItem = Split(arrItem(i), ",")(2)
        Set tplItem = tplGroup.Items.Add(lngId, strItem, xtpTaskItemTypeLink, lngImg)
        If i = lngUbound Then tplItem.SetMargins 0, 0, 0, 0 '不然最后一个选中时的框框不能完全框住内容
        If i + 1 = lngSelectedItem Then tplItem.Selected = True: tplFunc.Tag = lngId
    Next
    
End Sub

Public Sub LocatePar(ByRef txtInput As TextBox, ByRef objForm As Form)
'功能：查找参数并定位和显示
        Dim ctlTmp  As Control, strName As String
        Dim strInput As String, strOldColor As String
        Dim i As Long, p As Long, blnFind As Boolean
        Dim lngStart As Long, lngCount As Long
        Dim objPicPar As PictureBox
        Dim objTarget As Object
         
        lngStart = Val(txtInput.Tag)
        If lngStart = 0 Then lngStart = 1
        strInput = "*" & Trim(txtInput.Text) & "*"
      
        For Each ctlTmp In objForm.Controls
            
            lngCount = lngCount + 1
            If lngCount > lngStart Then
                strName = TypeName(ctlTmp)
                Select Case strName
                Case "Label", "CheckBox", "OptionButton", "Frame"
                    If ctlTmp.Caption Like strInput Then
                        blnFind = True
                        txtInput.Tag = lngCount
                        
                        '暂时最多支持三级
                        If ctlTmp.Container.Name = "picPar" Then
                            Set objPicPar = ctlTmp.Container
                        Else
                            On Error Resume Next    '如果不放在容器中，控件可能没有这么多级容器
                            If ctlTmp.Container.Container.Name = "picPar" Then
                                Set objPicPar = ctlTmp.Container.Container
                            ElseIf ctlTmp.Container.Container.Container.Name = "picPar" Then
                                Set objPicPar = ctlTmp.Container.Container.Container
                            End If
                            On Error GoTo 0: Err.Clear
                        End If
                        
                        If Not objPicPar Is Nothing Then
                            If objPicPar.Visible = False Then
                                For Each objTarget In objForm.picPar
                                    If objTarget Is objPicPar Then
                                        objTarget.Visible = True
                                        Call objForm.LocateFuncItem(objTarget.Index)
                                    Else
                                        objTarget.Visible = False
                                    End If
                                Next
                                objForm.Refresh
                            End If
                        End If
                        strOldColor = ctlTmp.ForeColor
                        
                        ctlTmp.ForeColor = vbRed
                        ctlTmp.Refresh
                        Call OS.Wait(400)
                        ctlTmp.ForeColor = &H80000012
                        ctlTmp.Refresh
                        Call OS.Wait(200)
                        ctlTmp.ForeColor = vbRed
                        ctlTmp.Refresh
                        Call OS.Wait(400)
                        ctlTmp.ForeColor = &H80000012
                        ctlTmp.Refresh
                        Call OS.Wait(200)
                        ctlTmp.ForeColor = vbRed
                        ctlTmp.Refresh
                        Call OS.Wait(400)
                        
                        ctlTmp.ForeColor = strOldColor
                        ctlTmp.Refresh
                        Exit For
                    End If
                End Select
            End If
        Next
        
        If blnFind = False Then
            If lngStart = 1 Then
                MsgBox "没有找到匹配的参数，请检查输入的内容。", vbInformation, "参数查找"
            Else
                MsgBox "全部找完了，后面没有了。", vbInformation, "参数查找"
                txtInput.Tag = ""
            End If
            
            txtInput.SelStart = 0
            txtInput.SelLength = Len(txtInput.Text)
            If txtInput.Enabled Then txtInput.SetFocus
        End If
End Sub


Public Sub EnterNextCell(ByRef vsobj As VSFlexGrid)
'功能：输框定位到下一个
    With vsobj
        If .Col + 1 > .Cols - 1 Then
            If .Row + 1 > .Rows - 1 Then .AddItem ""
            .Row = .Row + 1: .Col = .FixedCols
        Else
            .Col = .Col + 1
        End If
        '如果是隐藏行则递归再定位到下一个位置
        If .ColHidden(.Col) = True Then Call EnterNextCell(vsobj)
        .ShowCell .Row, .Col
    End With
End Sub

Public Function GetPar(ByRef rsPar As ADODB.Recordset, Optional ByVal strModules As String) As ADODB.Recordset
'功能：读取系统参数和指定的模块参数，返回记录集
'参数：模块号串
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    strSQL = "Select ID,参数名,参数号,Nvl(参数值,缺省值) as 参数值,NVL(部门, 0) As 部门,Nvl(私有, 0) 私有, NVL(本机, 0) 本机,影响控制说明,关联说明,适用说明,警告说明,Decode(警告说明,Null,0,1) as 是否关键参数,Nvl(模块,0) as 模块 " & vbCrLf & _
            "From Zlparameters Where 系统 = " & glngSys & "  And Nvl(性质,0) = 0 And " & _
            IIF(strModules = "", "模块 Is Null", "(模块 Is Null Or 模块 In(" & strModules & "))")
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "读取系统参数")
    Set rsPar = zlDatabase.CopyNewRec(rsTmp, False, "", Array("参数新值", adVarChar, 4000, Empty, "修改状态", adInteger, 1, Empty, _
                "控件名称", adVarChar, 50, Empty, "控件数组序号", adInteger, 3, Empty, "控件标识", adVarChar, 50, Empty, "ErrType", adInteger, 1, Empty))
    '标记私有本机参数
    Call rec.Update(rsPar, "(私有=1 And 部门=0) OR (本机=1 And 部门=0)", "ErrType", PET_不能设置)
    Set GetPar = rsTmp
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub SetParToControl(ByVal strPar As String, ByRef rsPar As ADODB.Recordset, ByRef arrObj As Variant, Optional ByVal bytMode As Byte = 0)
'功能：设置参数值到常用几类控件,并且在rsPar中建立参数与控件名称及数组序号的关联
'参数：strPar    -模块号1:参数号1(或参数名1):控件序号1,模块号2:参数号2(或参数名2):控件序号2,......，当arrObj为对象数组时传空
'      rsPar    -从数据库读取的参数记录集
'      arrObj   -支持Checkbox,ComboBox,UpDown,OptionButton,ListBox,TextBox控件数组(要求必须是数组)
'               -如果是对象数组，则格式为：模块1,参数名1,控件对象1,模块2,参数名2,控件对象2,......
'      bytMode- ListBox的ItemData取值模式：0-用Chr转换无分隔，1-直接用逗号分隔(*表示全部匹配),2-List(文本),3-匹配的不勾选,4-逗号分隔(包括全选)
'               ComboBox的取值模式：0-取ListIndex,1-取ItemData,2-val(List(i)),3-List(i)文本比较
'               OPtionButton 0-按参数保存Index，1-按参数第一位数为Index
    Dim strMsg As String, strErr As String
    Dim arrPar As Variant, i As Long, j As Long
    Dim lngModule As Long, lngPar As Long, strParName As String, lngObjIndex As Long
    Dim strType As String, strCtrlName As String
    Dim objTmp As Object
    
    On Error Resume Next
    
    If IsArray(arrObj) Then     'OptionButton数组控件的数组
        For i = 0 To UBound(arrObj) Step 3
            lngModule = arrObj(i)
            strParName = arrObj(i + 1)
            If IsNumeric(strParName) Then
                lngPar = Val(strParName): strParName = ""
            Else
                lngPar = 0
            End If
            Set objTmp = arrObj(i + 2)
            
            rsPar.Filter = IIF(strParName <> "", "参数名='" & strParName & "'", "参数号=" & lngPar) & " And 模块 = " & lngModule
            strType = TypeName(objTmp(0))
            strCtrlName = objTmp(0).Name
            If Err.Number <> 0 Then Err.Clear
            If rsPar.RecordCount > 0 Then
                If strType = "OptionButton" Then
                    rsPar!控件名称 = strCtrlName
                    If bytMode = 0 Then
                        objTmp(Val("" & rsPar!参数值)).value = True
                    ElseIf bytMode = 1 Then
                        If "" & rsPar!参数值 <> "" Then
                            objTmp(Val(Mid("" & rsPar!参数值, 1, 1))).value = True
                        Else
                            objTmp(0).value = True
                        End If
                    End If
                    If Err.Number <> 0 Then
                        Err.Clear
                        If Val(rsPar!ErrType & "") = 0 Then rsPar!ErrType = PET_值超限 '值范围不正确
                    End If
                End If
                rsPar!控件数组序号 = 0
                rsPar.Update
            Else
                '增加丢失的参数
                rsPar.AddNew Array("ID", "模块", "参数号", "参数名", "控件名称", "控件数组序号", "ErrType"), Array(-1, lngModule, lngPar, strParName, IIF(strType = "OptionButton", strCtrlName, Null), 0, PET_参数丢失)
            End If
        Next
    Else
        arrPar = Split(strPar, ",")
        strType = TypeName(arrObj(0))
        
        For i = 0 To UBound(arrPar)
            lngModule = Split(arrPar(i), ":")(0)
            strParName = Split(arrPar(i), ":")(1)
            If IsNumeric(strParName) Then
                lngPar = Val(strParName): strParName = ""
            Else
                lngPar = 0
            End If
            lngObjIndex = Split(arrPar(i), ":")(2)
            strCtrlName = arrObj(lngObjIndex).Name
            If strType = "UpDown" Then strCtrlName = "txtUD"
            rsPar.Filter = IIF(strParName <> "", "参数名='" & strParName & "'", "参数号=" & lngPar) & " And 模块 = " & lngModule
            If rsPar.RecordCount > 0 Then
                rsPar!控件名称 = strCtrlName
                Select Case strType
                    Case "CheckBox"
                        arrObj(lngObjIndex).value = IIF(Val("" & rsPar!参数值) <> 0, 1, 0)
                    Case "ComboBox"
                        If bytMode = 0 Then
                            arrObj(lngObjIndex).ListIndex = Val("" & rsPar!参数值)
                        Else
                            With arrObj(lngObjIndex)
                                For j = 0 To .ListCount - 1
                                    If bytMode = 1 Then
                                        If .ItemData(j) = Val("" & rsPar!参数值) Then
                                            .ListIndex = j
                                            Exit For
                                        End If
                                    ElseIf bytMode = 3 Then '文本比较
                                        If .List(j) = NVL(rsPar!参数值) Then
                                            .ListIndex = j: Exit For
                                        End If
                                    Else
                                        If Val(.List(j)) = Val("" & rsPar!参数值) Then
                                            .ListIndex = j
                                            Exit For
                                        End If
                                    End If
                                Next
                                If .ListCount > 0 And j > .ListCount - 1 Then .ListIndex = 0
                            End With
                        End If
                        arrObj(lngObjIndex).Tag = bytMode
                    Case "UpDown"
                        arrObj(lngObjIndex).value = rsPar!参数值
                    Case "OptionButton"
                        arrObj(Val(rsPar!参数值)).value = True
                        lngObjIndex = 0  '数组号固定存储为0
                    Case "TextBox"
                        arrObj(lngObjIndex).Text = NVL(rsPar!参数值)
                    Case "ListBox"
                        For j = 0 To arrObj(lngObjIndex).ListCount - 1
                            If bytMode = 0 Then
                                If InStr("" & rsPar!参数值, Chr(arrObj(lngObjIndex).ItemData(j))) > 0 Then
                                    arrObj(lngObjIndex).Selected(j) = True
                                End If
                            ElseIf bytMode = 1 Then
                                If "" & rsPar!参数值 = "*" Or InStr("," & rsPar!参数值 & ",", "," & arrObj(lngObjIndex).ItemData(j) & ",") > 0 Then
                                    arrObj(lngObjIndex).Selected(j) = True
                                End If
                            ElseIf bytMode = 3 Then
                                If InStr("" & rsPar!参数值, arrObj(lngObjIndex).ItemData(j)) = 0 Then
                                    arrObj(lngObjIndex).Selected(j) = True
                                End If
                            ElseIf bytMode = 4 Then
                                If InStr("," & rsPar!参数值 & ",", "," & arrObj(lngObjIndex).ItemData(j) & ",") > 0 Then
                                    arrObj(lngObjIndex).Selected(j) = True
                                End If
                            Else
                                If InStr("," & rsPar!参数值 & ",", "," & arrObj(lngObjIndex).List(j) & ",") > 0 Then
                                    arrObj(lngObjIndex).Selected(j) = True
                                End If
                            End If
                        Next
                        arrObj(lngObjIndex).Tag = bytMode
                    Case Else
                        If Val(rsPar!ErrType & "") = 0 Then rsPar!ErrType = PET_值超限 '值范围不正确
                End Select
                
                rsPar!控件数组序号 = lngObjIndex
                rsPar.Update
                If Err.Number <> 0 Then
                    Err.Clear
                    If Val(rsPar!ErrType & "") = 0 Then rsPar!ErrType = PET_值超限 '值范围不正确
                End If
            Else
                '增加丢失的参数
                rsPar.AddNew Array("ID", "模块", "参数号", "参数名", "控件名称", "控件数组序号", "ErrType"), Array(-1, lngModule, lngPar, strParName, strCtrlName, IIF(strType = "OptionButton", 0, lngObjIndex), PET_参数丢失)
            End If
        Next
    End If
'    rsPar.Filter = ""
End Sub

Public Sub SetParRelation(ByRef arrObj As Variant, ByVal lngObjIndex As Long, ByRef rsPar As ADODB.Recordset, _
                        Optional ByVal varPar As Variant, Optional ByVal lngModule As Long, _
                        Optional ByVal strObjTag As String, Optional strGridCol As String = "", _
                        Optional ByVal blnNotClearIndex As Boolean = False)
'功能：设置参数与控件的关联，以便控件悬浮提示时根据当前控件来查找参数来显示说明信息，以及用于关键参数的警告提示
'参数：varPar   -参数号过参数名，当值为0或空时，更新当前位置记录
'      lngModule-模块号，当值为0时，表示系统参数
'      strGridCol-绑定报表的列
'      blnNotClearIndex-不清除索引值（即变量:lngObjIndex(控件数组序号)）
    Dim strType As String, strObjName As String
    Dim lngPar As Long, strParName As String
    If TypeName(varPar) <> "Error" Then
        strParName = varPar & ""
        If IsNumeric(strParName) Then
            lngPar = Val(varPar)
            strParName = ""
        End If
    End If
    If lngPar <> 0 Or strParName <> "" Then
        rsPar.Filter = IIF(strParName <> "", "参数名='" & strParName & "'", "参数号=" & lngPar) & " And 模块 = " & lngModule
    End If
    
    strType = TypeName(arrObj)
    If strType = "Object" Then  '控件数组
        strObjName = arrObj(lngObjIndex).Name
        strType = TypeName(arrObj(lngObjIndex))
        If strType = "OptionButton" Then lngObjIndex = 0    '存入的时候固定是0，这里强制指定为0，不管传入值
    Else
        strObjName = arrObj.Name
        If blnNotClearIndex = False Then lngObjIndex = 0
    End If
    
    rsPar!控件名称 = strObjName & strGridCol
    rsPar!控件数组序号 = lngObjIndex
    rsPar!控件标识 = strObjTag
    rsPar.Update
End Sub


Public Sub SetParChange(ByRef arrObj As Variant, ByVal lngObjIndex As Long, ByRef rsPar As ADODB.Recordset, _
                        Optional ByVal blnValue As Boolean, _
                        Optional ByVal strValue As String, Optional strGridCols As String)
'功能：参数变化时，设置新值及修改状态
'参数：blnValue-指定参控件对应的参数值
'      strValue-如果是组合参数或特殊处理的参数,传入参数值（无法直接通过控件取值）
'      strGridCols-绑定的列(多个用逗号分离)
    Dim blnDo As Boolean
    Dim i As Long, strType As String
    Dim str类别 As String, bln全选 As Boolean
    Dim objTmp As Variant, varTemp As Variant, intCol As Integer
    Dim bytMode As Byte
'       ListBox的ItemData取值模式：0-用Chr转换无分隔，1-直接用逗号分隔(*表示全部匹配),2-List(文本)
'       ComboBox的取值模式：0-取ListIndex,1-取ItemData,2-val(List),3-List
    
    strType = TypeName(arrObj)
    
    If strType = "Object" Then  '控件数组
        strType = TypeName(arrObj(lngObjIndex))
        Set objTmp = arrObj(lngObjIndex)
    Else
        Set objTmp = arrObj
    End If
    If strType = "OptionButton" Then lngObjIndex = 0    '存入的时候固定是0，这里强制指定为0，不管传入值
    
    rsPar.Filter = "控件名称 = '" & objTmp.Name & strGridCols & "' And 控件数组序号=" & lngObjIndex
    If rsPar.RecordCount > 0 Then
        blnDo = True
        If blnValue Then
            rsPar!参数新值 = strValue
        Else
            Select Case strType
                Case "CheckBox"
                    rsPar!参数新值 = objTmp.value
                Case "ComboBox"
                    bytMode = Val(objTmp.Tag)
                    If bytMode = 0 Then
                        rsPar!参数新值 = objTmp.ListIndex
                    ElseIf bytMode = 1 Then
                        rsPar!参数新值 = objTmp.ItemData(objTmp.ListIndex)
                    ElseIf bytMode = 2 Then
                        rsPar!参数新值 = Val(objTmp.List(objTmp.ListIndex))
                    Else
                        rsPar!参数新值 = objTmp.List(objTmp.ListIndex)
                    End If
                Case "TextBox"  'UpDown关联的txtUD
                    rsPar!参数新值 = objTmp.Text
                Case "OptionButton"
                    For i = 0 To arrObj.UBound
                        If arrObj(i).value Then Exit For
                    Next
                    rsPar!参数新值 = i
                Case "ListBox"
                    bln全选 = True
                    bytMode = Val(objTmp.Tag)
                    For i = 0 To objTmp.ListCount - 1
                        If objTmp.Selected(i) Then
                            If bytMode = 0 Then
                                str类别 = str类别 & Chr(objTmp.ItemData(i))
                            ElseIf bytMode = 1 Then
                                str类别 = str类别 & "," & objTmp.ItemData(i)
                            ElseIf bytMode = 3 Then
                                '相反
                            ElseIf bytMode = 4 Then
                                str类别 = str类别 & "," & objTmp.ItemData(i)
                            Else
                                str类别 = str类别 & "," & objTmp.List(i)
                            End If
                        Else
                            If bytMode = 3 Then str类别 = str类别 & "," & objTmp.ItemData(i)
                            bln全选 = False
                        End If
                    Next
                    If bytMode = 1 Then
                        str类别 = IIF(bln全选, "*", Mid(str类别, 2))
                    ElseIf bytMode = 2 Then
                        str类别 = Mid(str类别, 2)
                    ElseIf bytMode = 3 Then
                        str类别 = Mid(str类别, 2)
                    ElseIf bytMode = 4 Then
                        str类别 = Mid(str类别, 2)
                    End If
                    
                    rsPar!参数新值 = str类别
                Case Else
                    blnDo = False
            End Select
        End If
        If blnDo Then
            rsPar.Update
            
            If "" & rsPar!参数新值 <> "" & rsPar!参数值 Then
                rsPar!修改状态 = 1
                If rsPar!是否关键参数 = 1 Then Call MsgBox("提醒：" & rsPar!警告说明, vbExclamation, "警告")
            Else
                rsPar!修改状态 = 0
            End If
            rsPar.Update
            
            Select Case strType
                Case "CheckBox", "ComboBox", "TextBox", "ListBox", "ListView"
                    objTmp.ForeColor = IIF(Val("" & rsPar!修改状态) = 1, &HC0&, &H0&)             '修改后用朱红色前景色标识
                Case "VSFlexGrid"
                    If strGridCols <> "" Then
                        varTemp = Split(strGridCols, ",")
                        For i = 0 To UBound(varTemp)
                            intCol = Val(varTemp(i))
                            objTmp.Cell(flexcpForeColor, objTmp.FixedRows, intCol, objTmp.Rows - 1, intCol) = IIF(Val("" & rsPar!修改状态) = 1, &HC0&, &H0&)
                        Next
                    Else
                        objTmp.ForeColor = IIF(Val("" & rsPar!修改状态) = 1, &HC0&, &H0&)             '修改后用朱红色前景色标识
                    End If
                Case "OptionButton"
                    For i = arrObj.LBound To arrObj.UBound
                        On Error Resume Next
                        If i = objTmp.Index Then
                            arrObj(i).ForeColor = IIF(Val("" & rsPar!修改状态) = 1, &HC0&, &H0&)
                        Else
                            arrObj(i).ForeColor = &H0& '其他的恢复黑色
                        End If
                        If Err.Number <> 0 Then Err.Clear
                        On Error GoTo 0
                    Next
            End Select
        End If
    End If
End Sub

Public Sub ShowErrParasMsg(ByRef objFrmMe As Object, ByRef rsPar As ADODB.Recordset)
'功能：错误参数提示，并将错误参数进行禁用等标识
'参数：objFrmMe=窗体
'         rsPar=参数记录集
'说明：该函数在参数加载完成后可以调用，只需调用一次即可
    Dim arrObject As Variant, arrTmp As Variant
    Dim objTmp As Object
    Dim strType As String, strCtrlName As String, blnArray As Boolean
    Dim strMsg As String, strTmp As String, petCurType As ParaErrType
    Dim intCount As Integer
    
    On Error GoTo errH
    '设置禁用颜色
    rsPar.Filter = "ErrType<>Null And ErrType<>" & PET_值超限
    Do While Not rsPar.EOF
        strCtrlName = rsPar!控件名称 & ""
        blnArray = False
        On Error Resume Next
        If strCtrlName <> "" Then
            Set arrObject = Nothing: Set objTmp = Nothing
            Set arrObject = CallByName(objFrmMe, strCtrlName, VbGet)
            If Err.Number <> 0 Then Err.Clear
            strType = TypeName(arrObject)
            If TypeName(arrObject) = "Object" Then
                blnArray = True
                For Each objTmp In arrObject
                    strType = TypeName(objTmp)
                    Exit For
                Next
            End If
            If strType <> "Empty" And strType <> "Nothing" Then
                If blnArray Then
                    Set objTmp = arrObject(Val(rsPar!控件数组序号 & ""))
                Else
                    Set objTmp = arrObject
                End If
                Select Case strType
                    Case "OptionButton"
                        For Each objTmp In arrObject
                            objTmp.ForeColor = &H808080
                            objTmp.Enabled = False
                        Next
                    Case "TextBox", "ComboBox"
                        objTmp.ForeColor = &H808080
                        objTmp.Locked = True
                        If strCtrlName = "txtUD" Then 'ud(UpDown)控件设置
                            Set arrTmp = CallByName(objFrmMe, "ud", VbGet)
                            If Err.Number <> 0 Then Err.Clear
                            Set objTmp = arrTmp(Val(rsPar!控件数组序号 & ""))
                            objTmp.ForeColor = &H808080
                            objTmp.Enabled = False
                        End If
                    Case Else
                        objTmp.ForeColor = &H808080
                        objTmp.Enabled = False
                End Select
            End If
        End If
        rsPar.MoveNext
    Loop
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo errH
    '消息提示：
    '不提示与控件无关联的本机私有参数
    rsPar.Filter = "(ErrType=" & PET_不能设置 & " And 控件名称<>Null ) OR (ErrType<>Null And ErrType<>" & PET_不能设置 & ")"
    rsPar.Sort = "ErrType,模块,参数号,参数名"
    petCurType = PET_正常
    strMsg = "": strTmp = ""
    Do While Not rsPar.EOF
        If petCurType <> Val(rsPar!ErrType) Then
            petCurType = Val(rsPar!ErrType): intCount = 0
            strMsg = strMsg & IIF(strMsg = "", "", vbNewLine) & strTmp
            strTmp = Decode(petCurType, PET_参数丢失, "以下参数未能正常读取，可能是缺少这些参数数据，请检查处理！", _
                                                            PET_不能设置, "以下参数由于类型变更为本机或私有参数，不能在此处设置，请到管理工具设置！", _
                                                            PET_值超限, "以下参数的值超过可用值范围，请检查处理！", "")
        End If
        strTmp = strTmp & IIF(intCount Mod 2 = 0, vbNewLine, ",  ") & IIF(rsPar!模块 = 0, "系统参数:", rsPar!模块 & "模块参数:") & IIF(rsPar!参数名 & "" <> "", rsPar!参数名, rsPar!参数号)
        intCount = intCount + 1
        rsPar.MoveNext
    Loop
    If strTmp <> "" Then
        strMsg = strMsg & IIF(strMsg = "", "", vbNewLine) & strTmp
    End If
    If strMsg <> "" Then
        MsgBox strMsg, vbExclamation, "注意"
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
End Sub

Public Function CheckParChanged(ByRef arrObj As Variant, ByVal lngObjIndex As Long, ByRef rsPar As ADODB.Recordset) As Boolean
'功能：根据指定的控件及数组号，判断对应的参数值是否改变
    Dim strType As String
    
    strType = TypeName(arrObj)
    If strType = "Object" Then  '控件数组
        strType = TypeName(arrObj(lngObjIndex))
        If strType = "OptionButton" Then lngObjIndex = 0    '存入的时候固定是0，这里强制指定为0，不管传入值
        
        rsPar.Filter = "控件名称='" & arrObj(lngObjIndex).Name & "' And 控件数组序号=" & lngObjIndex
    Else
        rsPar.Filter = "控件名称='" & arrObj.Name & "' And 控件数组序号=0"
    End If
    
    If rsPar.RecordCount > 0 Then
        CheckParChanged = (Val("" & rsPar!修改状态) = 1)
    End If
End Function


Public Function GetParOriginalValue(ByRef arrObj As Variant, ByVal lngObjIndex As Long, ByRef rsPar As ADODB.Recordset) As String
'功能：根据指定的控件及数组号，返回参数原始值
    Dim strType As String
    
    strType = TypeName(arrObj)
    If strType = "Object" Then  '控件数组
        strType = TypeName(arrObj(lngObjIndex))
        If strType = "OptionButton" Then lngObjIndex = 0    '存入的时候固定是0，这里强制指定为0，不管传入值
            
        rsPar.Filter = "控件名称='" & arrObj(lngObjIndex).Name & "' And 控件数组序号=" & lngObjIndex
    Else
        rsPar.Filter = "控件名称='" & arrObj.Name & "' And 控件数组序号=0"
    End If
    If rsPar.RecordCount > 0 Then
        GetParOriginalValue = rsPar!参数值
    End If
End Function

Public Function SavePar(ByRef rsPar As ADODB.Recordset, ByRef frmParent As Form) As Boolean
'功能：保存修改过的参数
    Dim strPars As String, strPar As String
    
    With rsPar
        '只进行没有错误的参数与值超过范围的参数设置
        .Filter = "(修改状态=1 ANd ErrType =Null) OR  (修改状态=1 And ErrType=" & PET_值超限 & ")"
        Do While Not .EOF
            If InStr(!参数新值, gstrParSplit1) > 0 Or InStr(!参数新值, gstrParSplit2) > 0 Then
                MsgBox "模块" & !模块 & "的参数[" & !参数名 & "]含有非法的" & gstrParSplit1 & "或" & gstrParSplit2 & "，不允许保存!" & vbCrLf & _
                    "参数值:" & !参数新值, vbExclamation, "错误"
                Exit Function
            End If
            '由于目前参数值中包含了:,|等字符，所以选取^#为分隔符
            strPar = !模块 & gstrParSplit1 & !参数号 & gstrParSplit1 & !参数新值
            
            '有多个控件对应一个参数的情况，需去除重复
            If InStr(strPars, strPar) = 0 Then strPars = strPars & gstrParSplit2 & strPar
            .MoveNext
        Loop
    End With
    strPars = Mid(strPars, 2)
    
    If strPars <> "" Then
        On Error GoTo ErrHandle
        gstrSQL = "zl_Parameters_Update_Batch(" & glngSys & ",'" & strPars & "','" & gstrUserName & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "参数保存")
        
        '如果修改了关键参数，则弹出界面输入原因并保存
        rsPar.Filter = "(是否关键参数=1 And 修改状态=1 ANd ErrType =Null) OR  (是否关键参数=1 And 修改状态=1 And ErrType=" & PET_值超限 & ")"
        If rsPar.RecordCount > 0 Then
            Call frmParReason.ShowMe(frmParent, rsPar)
        End If
    End If
    SavePar = True
    Call zlDatabase.ClearParaCache '清空参数缓存
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Call zlDatabase.ClearParaCache '清空参数缓存
End Function

Public Function GetFuncID(ByVal strName As String, ByRef arrFunc As Variant) As Long
'功能：根据名称返回二级分类ID
'参数：arrFunc-多个二级分类名称数组，以分号分隔,以逗号分隔图标ID、容器数组及二级分类名称,例：401,1,门诊划价管理;412,2,病人收费管理;......
'返回：功能ID
    Dim i As Long, j As Long
    Dim arrTmp As Variant
    
    For i = 0 To UBound(arrFunc)
        arrTmp = Split(arrFunc(i), ";")
        For j = 0 To UBound(arrTmp)
            If strName = Split(arrTmp(j), ",")(2) Then
                GetFuncID = Split(arrTmp(j), ",")(1)
                Exit Function
            End If
        Next
    Next
End Function

Public Sub SetParTip(ByRef arrObj As Variant, ByVal lngObjIndex As Long, ByRef rsPar As ADODB.Recordset, _
    Optional ByVal strObjTag As String, Optional ByVal objOtherControl As Object, _
    Optional ByVal strGridCol As String)
'功能：根据控件及数组号返回组织好的参数提示文本
'objOtherControl：指定在其他控件上显示提示文本
'strGridCol-绑定的列
    Dim strTip As String
    Dim strType As String
    Dim blnArray As Boolean
    Dim petCur As ParaErrType
    strType = TypeName(arrObj)
    If strType = "Object" Then  '控件数组
        blnArray = True
        strType = TypeName(arrObj(lngObjIndex))
        If strType = "OptionButton" Then  '存入的时候固定是0，这里强制指定为0，不管传入值
            rsPar.Filter = "控件名称='" & arrObj(lngObjIndex).Name & strGridCol & "' And 控件数组序号=0"
        Else
            rsPar.Filter = "控件名称='" & arrObj(lngObjIndex).Name & strGridCol & "' And 控件数组序号=" & lngObjIndex & IIF(strObjTag <> "", " And 控件标识='" & strObjTag & "'", "")
        End If
    Else
        rsPar.Filter = "控件名称='" & arrObj.Name & strGridCol & "' And 控件数组序号=0" & " And 控件标识='" & strObjTag & "'"
    End If
    
    If rsPar.RecordCount > 0 Then
        petCur = Val(rsPar!ErrType & "")
        strTip = IIF(rsPar!模块 = 0, "系统全局参数", "模块号：" & rsPar!模块) & "，参数号：" & rsPar!参数号
        
        If petCur = PET_不能设置 Or petCur = PET_参数丢失 Then
            strTip = strTip & vbCrLf & "参数设置警告|" & IIF(petCur = PET_不能设置, "该参数当前类型为本机或私有参数，不能在此处设置，请到管理工具设置。", "该参数未能正常读取，可能是缺少该参数数据。")
        End If
        strTip = strTip & vbCrLf & "影响控制说明|" & rsPar!影响控制说明
        If Not IsNull(rsPar!关联说明) Then strTip = strTip & vbCrLf & "关联说明|" & rsPar!关联说明
        If Not IsNull(rsPar!适用说明) Then strTip = strTip & vbCrLf & "适用说明|" & rsPar!适用说明
        If Not IsNull(rsPar!警告说明) Then strTip = strTip & vbCrLf & "警告说明|" & rsPar!警告说明
    End If
    
    If strTip <> "" Then
        If Not objOtherControl Is Nothing Then
            Call zlCommFun.ShowTipInfo(objOtherControl.hwnd, strTip, True, True, 8800)
        ElseIf blnArray Then
            Call zlCommFun.ShowTipInfo(arrObj(lngObjIndex).hwnd, strTip, True, True, 8800)
        Else
            Call zlCommFun.ShowTipInfo(arrObj.hwnd, strTip, True, True, 8800)
        End If
    End If
End Sub

Public Sub SetPrompt(ByRef lblPrompt As Label, ByVal strPrompt As String)
'功能：设置提示信息，稍后自动消失
    lblPrompt.Caption = strPrompt
    lblPrompt.Refresh
    Call OS.Wait(2500)
    lblPrompt.Caption = ""
End Sub


Public Sub SetVsfEditable(ByRef vsf As VSFlexGrid, ByVal blnEdit As Boolean)
'功能：设置表格控件的可用性及外观
    With vsf
        .Enabled = blnEdit
        .Editable = IIF(blnEdit, flexEDKbdMouse, flexEDNone)
        .ForeColor = IIF(blnEdit, vsf.Container.ForeColor, &H808080)
        .BackColor = IIF(blnEdit, &H80000005, vsf.Container.BackColor)
    End With
End Sub

Public Sub SetLstSelected(ByRef lst As ListBox, ByVal blnSel As Boolean)
'功能：全选或全消ListBox项目，保持位置不变
    Dim i As Long, Y As Long
    
    With lst
        Y = .ListIndex
        For i = 0 To .ListCount - 1
            .Selected(i) = blnSel    '将触发lst_ItemCheck事件
        Next
        .ListIndex = Y
    End With
End Sub




