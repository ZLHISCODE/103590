VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisResultItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################
'接口种类： 统计分析
'实现功能：
'编制人员：
'编制时间：
'接口类名：
'接口方法： 1.InitElement
'           2.GetElementData
'           3.SetElementProperty
'           4.ShowElement
'           5.GetElementObject
'接口属性： 1.CurrentRect
'           2.CurrentFont
'######################################################################################################################

Private mrsParameter                As ADODB.Recordset
Private mbytApplyMode               As Byte                               '1-界面执行显示;2-打印输出
Private mobjRect                    As USERRECT
Private mobjFont                    As USERFONT
Private mfrmMain                    As Object
Private mclsLisPackage              As clsLisPackage
Private mclsLisReportData           As clsLisReportData
Private mclsLisReportLib            As clsLisReportLib
Private mclsDrawLisImage            As clsDrawLisImage
Private mlngGridRowHeight           As Single                               '单位：厘米
Private mlngGridColWidth(1 To 7)    As Single                              '单位：厘米
Private mintCount                   As Integer
Private mX0                         As Single
Private mY0                         As Single
Private mblnSingular                As Boolean
Private mSingular                   As Single
Private mblnEnzymeMarker            As Boolean              '是否酶标项目
'公共接口
'######################################################################################################################
Public Property Get CurrentRect() As USERRECT
    CurrentRect = mobjRect
End Property

Public Property Let CurrentRect(vData As USERRECT)
    mobjRect = vData
End Property

Public Property Get CurrentFont() As USERFONT
    CurrentFont = mobjFont
End Property

Public Property Let CurrentFont(vData As USERFONT)
    mobjFont = vData
End Property

Public Function InitElement(ByVal bytApplyMode As Byte, Optional frmMain As Object = Nothing) As Boolean
    '******************************************************************************************************************
    '功能：初始化组件，并传入数据库连接对象以及调用窗体对象
    '参数：bytApplyMode   应用场合;cnOracle 数据库连接;frMain 调用窗体对象
    '返回：初始化成功返回True,否则返回False
    '******************************************************************************************************************
    mbytApplyMode = bytApplyMode
    Set mfrmMain = frmMain
    
    Set mclsLisPackage = New clsLisPackage
    Set mclsLisReportData = New clsLisReportData
    Set mclsLisReportLib = New clsLisReportLib
    Set mclsDrawLisImage = New clsDrawLisImage
    Set gclsLisReportLib = New clsLisReportLib
    InitElement = True
    
End Function

Public Function SetElementParameter(ByVal rsParameter As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：设置组件属性
    '参数：rsProperty    属性记录集
    '返回：设置成功返回True,否则返回False
    '******************************************************************************************************************
    Set mrsParameter = rsParameter
    
'    SetElementParameter = mfrmElementExcept.SetParameter(rsParameter)
End Function

Public Function RefreshElement() As Boolean
    '******************************************************************************************************************
    '功能：根据属性查询并显示在界面上，只用于界面执行显示
    '参数：无
    '返回：打开成功返回True;否则返回False
    '******************************************************************************************************************
    
    If mbytApplyMode = 2 Then Exit Function
    
    RefreshElement = True
    
End Function

Public Function GetElementObject() As Object
    '******************************************************************************************************************
    '功能：获取用于显示组件内容的窗体对象，只用于界面执行显示
    '参数：无
    '返回：获取成功返回True;否则返回False
    '******************************************************************************************************************
    If mbytApplyMode = 2 Then Exit Function
    
    
    Set GetElementObject = Nothing
    
End Function

Public Function GetElementData() As Boolean
    '******************************************************************************************************************
    '功能：只用于打印预览输出
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objRect                 As USERRECT
    Dim objFont                 As USERFONT

    On Error GoTo errHand
    
    If mbytApplyMode = 1 Then Exit Function
    
    '1.获取当前要输出的起始位置(下面的代码不能删除)
    '------------------------------------------------------------------------------------------------------------------
    objRect = CurrentRect
    
    '2.获取当前的字体设置(下面的代码不能删除)
    '------------------------------------------------------------------------------------------------------------------
    objFont = CurrentFont

    '3.以下是自定义输出部份处理
    '------------------------------------------------------------------------------------------------------------------
    Dim rsItem              As ADODB.Recordset
    Dim intRows             As Integer
    Dim intColRows          As Variant
    Dim rsTmp               As ADODB.Recordset
    Dim rsGermResult        As ADODB.Recordset
    Dim rsResult            As ADODB.Recordset
    Dim rsAntibioticResult  As ADODB.Recordset
    Dim rsConclusion        As ADODB.Recordset
    Dim rsMultiline         As ADODB.Recordset      '
    Dim rsGermsCount        As ADODB.Recordset
    Dim rsEnzymeMarker      As ADODB.Recordset
    Dim rsAdvice            As ADODB.Recordset
    Dim varTmp              As Variant
    Dim strTmp              As String
    Dim strItemTmp          As String
    Dim str结论描述         As String
    Dim lngTmp              As Long
    Dim blnOutPutTitle      As Boolean
    Dim objSvrRect          As USERRECT
    Dim objSvrFont          As USERFONT
    Dim intCount            As Integer
    Dim strItemName         As String
    Dim lngMapHeight        As Long
    Dim blnPutOut           As Boolean
    Dim dblFontSize         As Double
    Dim blnMultiline        As Boolean              '是否多参考
    Dim blnNoGermTest       As Boolean
    Dim intMaxLine          As Integer
    Dim bln第二列           As Boolean
    Dim intItemCount        As Integer
    Dim intStartRow         As Integer
    Dim intEndRow           As Integer
    Dim blnNewPage          As Boolean
    Dim strImagePath        As String
    Dim btyType             As Integer
    Dim str微生物评语       As String
    Dim rsSignId            As ADODB.Recordset
    Dim rsSignPicture       As ADODB.Recordset
    Dim strSign             As String
    Dim intRowNum           As Integer
    Dim strView             As String
    Dim strViewRowNum       As String
    Dim blnAddRow           As Boolean
    
    mlngGridRowHeight = 0.6
    
    dblFontSize = objFont.Size
    '事先打开当前病人的所有体检项目
    Set rsTmp = mclsLisPackage.Get_RequestOrder(Val(mclsLisPackage.GetParameter(mrsParameter, "病人ID")), Val(mclsLisPackage.GetParameter(mrsParameter, "主页ID")), mclsLisPackage.GetParameter(mrsParameter, "申请开始时间"), mclsLisPackage.GetParameter(mrsParameter, "申请结束时间"))
    intColRows = Split("1;1;1;1;1;1;1;1", ";")
'    rsItem.Sort = "类别,分类"
    objRect.Y1 = objRect.Y0
    mintCount = 0
    intItemCount = 0
    intStartRow = 0
    intEndRow = 0
    intCount = 1
    blnNewPage = False
    
    If Not ChkRsState(rsTmp) Then
        Set rsItem = CopyRecordStruct(rsTmp, , True)
        rsItem.Fields.Append "类别顺序", adBigInt
        rsItem.Fields.Append "标本顺序", adBigInt
        rsItem.Fields.Append "项目顺序", adBigInt
        rsItem.Open
        Call CopyRecordData(rsTmp, rsItem)
        rsItem.MoveFirst
        
        strTmp = mclsLisPackage.GetParameter(mrsParameter, "类别顺序")
        If strTmp <> "" Then
            varTmp = Split(strTmp, ",")
            For intCount = 0 To UBound(varTmp)
                rsItem.Filter = ""
                rsItem.Filter = "分类='" & Split(varTmp(intCount), "#")(0) & "'"
                If rsItem.RecordCount > 0 Then
                    Do While Not rsItem.EOF
                        rsItem("类别顺序").Value = Split(varTmp(intCount), "#")(1)
                        rsItem.MoveNext
                    Loop
                End If
            Next
        Else
            rsItem.Filter = ""
            If rsItem.RecordCount > 0 Then
                intCount = 0
                Do While Not rsItem.EOF
                    rsItem("类别顺序").Value = intCount
                    intCount = intCount + 1
                    rsItem.MoveNext
                Loop
            End If
        End If
        
        rsItem.Filter = ""
        rsItem.MoveFirst
        
        strTmp = mclsLisPackage.GetParameter(mrsParameter, "标本顺序")
        
        If strTmp <> "" Then
            varTmp = Split(strTmp, ",")
            For intCount = 0 To UBound(varTmp)
                rsItem.Filter = ""
                rsItem.Filter = "标本类型='" & Split(varTmp(intCount), "#")(0) & "'"
                If rsItem.RecordCount > 0 Then
                    Do While Not rsItem.EOF
                        rsItem("标本顺序").Value = Split(varTmp(intCount), "#")(1)
                        rsItem.MoveNext
                    Loop
                End If
            Next
        Else
            rsItem.Filter = ""
            If rsItem.RecordCount > 0 Then
                intCount = 0
                Do While Not rsItem.EOF
                    rsItem("标本顺序").Value = intCount
                    intCount = intCount + 1
                    rsItem.MoveNext
                Loop
            End If
        End If
        
        rsItem.Filter = ""
        rsItem.MoveFirst
        
        strTmp = mclsLisPackage.GetParameter(mrsParameter, "项目顺序")
        
        If strTmp <> "" Then
            varTmp = Split(strTmp, ",")
            For intCount = 0 To UBound(varTmp)
                rsItem.Filter = ""
                rsItem.Filter = "检验项目1='" & Split(varTmp(intCount), "#")(0) & "'"
                If rsItem.RecordCount > 0 Then
                    Do While Not rsItem.EOF
                        rsItem("项目顺序").Value = Split(varTmp(intCount), "#")(1)
                        rsItem.MoveNext
                    Loop
                End If
            Next
        Else
            rsItem.Filter = ""
            If rsItem.RecordCount > 0 Then
                intCount = 0
                Do While Not rsItem.EOF
                    rsItem("项目顺序").Value = intCount
                    intCount = intCount + 1
                    rsItem.MoveNext
                Loop
            End If
        End If
        
        rsItem.Filter = ""
        rsItem.Sort = "类别顺序,标本顺序,项目顺序"
        
        Do While Not rsItem.EOF
            intItemCount = intItemCount + 1
    '        intStartRow = intStartRow - 1
            mintCount = 0
            intStartRow = 0
            mX0 = 0
            mY0 = 0
            bln第二列 = False
            
            objFont.BackColor = USERCOLOR.白色
            objFont.ForeColor = USERCOLOR.黑色
            objFont.Underline = False
            gclsLisReportLib.Font = objFont
    
            objRect.X0 = gclsLisReportLib.Paper.BorderLeft
            objRect.Y0 = objRect.Y1 + gclsLisReportLib.GetTwipsY(0.2)
            objRect.X1 = 0
            objRect.Y1 = 0
            objRect.B0 = 0
    
            '先计算一下从此行开始，是否有数据行在本页，如果没有，则换页
            '--------------------------------------------------------------------------------------------------
            lngTmp = objRect.Y0 + gclsLisReportLib.GetTextHeight("高") + 60 + 3 * gclsLisReportLib.GetTwipsX(mlngGridRowHeight)
            If lngTmp > gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - gclsLisReportLib.Paper.PageFoot - gclsLisReportLib.Paper.SpaceBottom Then
                Call mclsLisReportData.NewPage
                objRect.Page = objRect.Page + 1
                objRect.Y0 = gclsLisReportLib.Paper.PageHead + gclsLisReportLib.Paper.BorderTop + gclsLisReportLib.GetTwipsY(0.2) '- gclsLisReportLib.GetTwipsY(0.2) '
                objRect.Y1 = 0
            End If
            '*******************************插入报告单表头******************************************************
            objFont.Name = "宋体"
            objFont.Size = 13
            objFont.Bold = True
            objRect.X0 = gclsLisReportLib.Paper.Width / 2 - 2340 'gclsLisReportLib.GetTextWidth(gstrHospital & "检验报告单")
            objRect.X1 = 0
            Call mclsLisReportData.InsertText(gstrHospital & "检验报告单", objRect, objFont)
    '        Call mclsLisReportData.AppendList(gstrHospital & "检验报告单", 2, objRect.Page)
            
            objFont.Size = 8
            objFont.Bold = False
        
            Call MakeDeviceName(objRect, objFont, rsItem)
            
            objRect.Y1 = objRect.Y1 + 30
            objFont.ForeColor = USERCOLOR.表格线色
            
            Call SetRect(objRect, gclsLisReportLib.Paper.BorderLeft, objRect.Y1, gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight, objRect.Y1)
    
            Call mclsLisReportData.InsertLine(objRect, objFont)
            '*******************************插入报告单信息及患者信息******************************************************
    
        
            objFont.Size = 9.5
            objFont.ForeColor = USERCOLOR.黑色
            objFont.Bold = False
            btyType = 1
            Call MakePatientTitle(objRect, objFont, btyType, rsItem, "")
            
            objRect.Y1 = objRect.Y1 + 30
            objFont.ForeColor = USERCOLOR.表格线色
            
            Call SetRect(objRect, gclsLisReportLib.Paper.BorderLeft, objRect.Y1, gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight, objRect.Y1)
            
            Call mclsLisReportData.InsertLine(objRect, objFont)
            
             '判断是否为阴性报告，如果阴性报告单就为纯文本报告
            Set rsGermsCount = mclsLisPackage.Get_LisGermsCount(Val(Nvl(rsItem!标本id)))
            blnNoGermTest = (Val(rsGermsCount!数量) = 0)
             '判断是否有多参考，多参考使用单列显示
            Set rsMultiline = mclsLisPackage.Get_LisMultiLineRefer(Val(Nvl(rsItem!标本id)))
            Set rsEnzymeMarker = mclsLisPackage.Get_LisEnzymeMarker(Val(Nvl(rsItem!标本id)))
            blnMultiline = (Val(rsMultiline!临床特诊) >= 1)
            If Not blnMultiline Then blnMultiline = (Val(rsEnzymeMarker!数量) >= 1)
            mblnEnzymeMarker = (Val(rsEnzymeMarker!数量) >= 1)
            
            If blnMultiline And blnNoGermTest Then
                If mblnEnzymeMarker Then
'                    mlngGridColWidth(1) = 0.5
                    mlngGridColWidth(1) = 1
                    mlngGridColWidth(2) = 7.5
                    mlngGridColWidth(3) = 4
                    mlngGridColWidth(4) = 2.5
                    mlngGridColWidth(5) = 2.3
                    mlngGridColWidth(6) = 1

                Else
'                    mlngGridColWidth(1) = 0.5
                    mlngGridColWidth(1) = 1
                    mlngGridColWidth(2) = 7.5
                    mlngGridColWidth(3) = 4
                    mlngGridColWidth(4) = 0.5
                    mlngGridColWidth(5) = 4.3
                    mlngGridColWidth(6) = 2
                End If
            Else
'                mlngGridColWidth(1) = 0.5
                mlngGridColWidth(1) = 1
                mlngGridColWidth(2) = 3.4
                mlngGridColWidth(3) = 1.5
                mlngGridColWidth(4) = 0.6
                mlngGridColWidth(5) = 2.4
                mlngGridColWidth(6) = 1.2
            End If
            '*******************************插入检验结果信息******************************************************
            If Val(Nvl(rsItem!微生物申请)) = 1 Then
                '判断是否为阴性报告，如果阴性报告单就为纯文本报告
                str微生物评语 = ""
                If blnNoGermTest Then
                    '阴性报告，为纯文本
                    Set rsResult = mclsLisPackage.Get_LisGermFreeResult(Val(Nvl(rsItem!标本id)))
                    If Not ChkRsState(rsResult) Then
                        str微生物评语 = Nvl(rsResult!阴性评语)
                        lngTmp = objRect.Y0 + 3 * gclsLisReportLib.GetTwipsX(mlngGridRowHeight)
                        If lngTmp > gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - gclsLisReportLib.Paper.PageFoot - gclsLisReportLib.Paper.SpaceBottom Then
                            Call mclsLisReportData.NewPage
                            objRect.Page = objRect.Page + 1
                            objRect.Y0 = gclsLisReportLib.Paper.PageHead + gclsLisReportLib.Paper.BorderTop + gclsLisReportLib.GetTwipsY(0.2)
                            objRect.Y1 = objRect.Y0
                        End If
                         '产生项目标题
                        objRect.Y1 = objRect.Y1 + 15
                        strItemName = "仪  器:" & rsItem!仪器 'Nvl(rsItem!检验项目)
                        Call MakeItemTitle(objRect, objFont, strItemName)
                        If gblnNewLis Then
                            strItemTmp = "正常菌[" & Nvl(rsResult!正常菌) & "],未检出[" & Nvl(rsResult!未检出) & "],正常菌群为[" & Nvl(rsResult!正常菌群) & "]。"
                        Else
                            strItemTmp = "正常菌[" & Nvl(rsResult!检验项目) & "],检验结果[" & Nvl(rsResult!检验结果) & "],结果描述[" & Nvl(rsResult!结果描述) & "]。"
                        End If
                        If strItemTmp <> "" Then
                            objFont.ForeColor = USERCOLOR.黑色
                            objFont.Bold = False
                            '    objFont.Size = 12
                            gclsLisReportLib.Font = objFont
                            
                            
                            objRect.R0 = gclsLisReportLib.GetTwipsX(0.15)
                            objRect.X0 = gclsLisReportLib.Paper.BorderLeft
                            objRect.Y0 = objRect.Y1 + gclsLisReportLib.GetTwipsX(0.2)
                            objRect.Y1 = objRect.Y0 + 2 * objRect.B0 + gclsLisReportLib.GetTextHeight("高") + gclsLisReportLib.GetTwipsX(0.15)
                            
                            objFont.BackColor = USERCOLOR.项目标题背景色
                            objRect.X0 = gclsLisReportLib.Paper.BorderLeft
                            objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
                            
                '            Call mclsLisReportData.InsertBackColor(objRect, objFont)
                            
                            objFont.ForeColor = USERCOLOR.黑色
                            objFont.BackColor = USERCOLOR.项目标题背景色
                            objRect.Y0 = objRect.Y0 + 30
                            objRect.Y1 = objRect.Y1
                            objFont.WordWarp = True
                            strItemTmp = Get_Text("检测结果:" & strItemTmp, gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderLeft - gclsLisReportLib.Paper.BorderRight, Abs(Int(-objRect.X0 / gclsLisReportLib.GetTextWidth(Space(1)))))
                            Call mclsLisReportData.InsertText(strItemTmp, objRect, objFont)
                        End If
                    End If
                Else
                    
                    '产生项目标题
                    objRect.Y1 = objRect.Y1 + 15
                    strItemName = "仪  器:" & rsItem!仪器 'Nvl(rsItem!检验项目)
                    Call MakeItemTitle(objRect, objFont, strItemName)
                    lngTmp = objRect.Y0 + 3 * gclsLisReportLib.GetTwipsX(mlngGridRowHeight)
                    If lngTmp > gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - gclsLisReportLib.Paper.PageFoot - gclsLisReportLib.Paper.SpaceBottom Then
                        Call mclsLisReportData.NewPage
                        objRect.Page = objRect.Page + 1
                        objRect.Y0 = gclsLisReportLib.Paper.PageHead + gclsLisReportLib.Paper.BorderTop + gclsLisReportLib.GetTwipsY(0.2)
                        objRect.Y1 = objRect.Y0
                    End If
                    '细菌结果
                    Set rsGermResult = mclsLisPackage.Get_LisGermsResult(Val(Nvl(rsItem!标本id)))
                    If Not ChkRsState(rsGermResult) Then
                        Do While Not rsGermResult.EOF
                            '抗生素结果
                            Set rsResult = mclsLisPackage.Get_LisAntisResult(Val(Nvl(rsItem!标本id)), Val(Nvl(rsGermResult!细菌ID)))
                            str微生物评语 = Nvl(rsResult!阳性评语)
                            If Not ChkRsState(rsResult) Then
                                '                    Call MakeBacilliResultData(objRect, objFont, rsItem, rsGermResult)
                                
                                '                        lngTmp = objRect.Y0 + 3 * gclsLisReportLib.GetTwipsX(mlngGridRowHeight)
                                '                        If lngTmp > gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - gclsLisReportLib.Paper.PageFoot - gclsLisReportLib.Paper.SpaceBottom Then
                                '                            Call mclsLisReportData.NewPage
                                '                            objRect.Page = objRect.Page + 1
                                '                            objRect.Y0 = gclsLisReportLib.Paper.PageHead + gclsLisReportLib.Paper.BorderTop + gclsLisReportLib.GetTwipsY(0.2)
                                '                            objRect.Y1 = objRect.Y0
                                '                        End If
                                '                         '产生细菌结果
                                '                        objFont.ForeColor = USERCOLOR.黑色
                                '                        objFont.BackColor = USERCOLOR.项目标题背景色
                                ''                        objRect.Y0 = objRect.Y0 + 400
                                '                        objRect.Y1 = objRect.Y1 + gclsLisReportLib.GetTwipsY(0.8)
                                '                        strItemName = "鉴定结果:" & Nvl(rsGermResult!中文名)
                                '                        Call mclsLisReportData.InsertText(strItemName, objRect, objFont)
                                '
                                '                        objRect.Y1 = objRect.Y1 + gclsLisReportLib.GetTextHeight("高")
                                '                        objFont.ForeColor = USERCOLOR.表格线色
                                '
                                '                        Call SetRect(objRect, gclsLisReportLib.Paper.BorderLeft, objRect.Y1, gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight, objRect.Y1)
                                '                        Call mclsLisReportData.InsertLine(objRect, objFont)
                                Call MakeBacilliResultData(objRect, objFont, rsGermResult)
                                '产生抗生素结果标题
                                Call MakeAntibioticResultTitle(objRect, objFont)
                                '产生抗生素结果
                                Call MakeAntibioticResultData(objRect, objFont, rsResult)
                                
                            End If
                            rsGermResult.MoveNext
                        Loop
                    End If
                End If
                '产生结果评语
                If Nvl(str微生物评语) <> "" Then
                    lngTmp = objRect.Y0 + 3 * gclsLisReportLib.GetTwipsX(mlngGridRowHeight)
                    If lngTmp > gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - gclsLisReportLib.Paper.PageFoot - gclsLisReportLib.Paper.SpaceBottom Then
                        Call mclsLisReportData.NewPage
                        objRect.Page = objRect.Page + 1
                        objRect.Y0 = gclsLisReportLib.Paper.PageHead + gclsLisReportLib.Paper.BorderTop + gclsLisReportLib.GetTwipsY(0.2)
                        objRect.Y1 = objRect.Y0
                    End If
                    objRect.Y1 = objRect.Y1 + 15
                    If str微生物评语 <> "" Then
                        objFont.ForeColor = USERCOLOR.黑色
                        objFont.Bold = False
                        '    objFont.Size = 12
                        gclsLisReportLib.Font = objFont
                        
                        
                        objRect.R0 = gclsLisReportLib.GetTwipsX(0.15)
                        objRect.X0 = gclsLisReportLib.Paper.BorderLeft
                        objRect.Y0 = objRect.Y1 + gclsLisReportLib.GetTwipsX(0.2)
                        objRect.Y1 = objRect.Y0 + 2 * objRect.B0 + gclsLisReportLib.GetTextHeight("高") + gclsLisReportLib.GetTwipsX(0.15)
                        
                        objFont.BackColor = USERCOLOR.项目标题背景色
                        objRect.X0 = gclsLisReportLib.Paper.BorderLeft
                        objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
                        
            '            Call mclsLisReportData.InsertBackColor(objRect, objFont)
                        
                        objFont.ForeColor = USERCOLOR.黑色
                        objFont.BackColor = USERCOLOR.项目标题背景色
                        objRect.Y0 = objRect.Y0 + 30
                        objRect.Y1 = objRect.Y1
                        objFont.WordWarp = True
                        strItemTmp = Get_Text("结果评语:" & str微生物评语, gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderLeft - gclsLisReportLib.Paper.BorderRight, Abs(Int(-objRect.X0 / gclsLisReportLib.GetTextWidth(Space(1)))))
                        Dim VarItemTmp As Variant
                        Dim i As Long
                        If InStr(1, strItemTmp, vbCrLf) > 0 Then
                            '判断评语中，有回车的情况，进行换行处理
                            VarItemTmp = Split(strItemTmp, vbCrLf)
                            For i = 0 To UBound(VarItemTmp)
                                If VarItemTmp(i) <> "" Then
                                    objRect.Y1 = objRect.Y1 + 15
                                    objFont.WordWarp = True
                                     Call MakeItemTitle(objRect, objFont, VarItemTmp(i))
                                End If
                            Next
                        Else
                            Call mclsLisReportData.InsertText(strItemTmp, objRect, objFont)
                        End If
                    End If
                End If
            Else
    
                Set rsResult = mclsLisPackage.Get_LisResult(Val(Nvl(rsItem!标本id)), blnMultiline)
                If Not ChkRsState(rsResult) Then
                   
                    lngTmp = objRect.Y0 + 3 * gclsLisReportLib.GetTwipsX(mlngGridRowHeight)
                    If lngTmp > gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - gclsLisReportLib.Paper.PageFoot - gclsLisReportLib.Paper.SpaceBottom Then
                        Call mclsLisReportData.NewPage
                        objRect.Page = objRect.Page + 1
                        objRect.Y0 = gclsLisReportLib.Paper.PageHead + gclsLisReportLib.Paper.BorderTop + gclsLisReportLib.GetTwipsY(0.2)
                        objRect.Y1 = objRect.Y0
'                        mintCount = 0
                    End If
                    blnAddRow = False
                    '产生项目标题
                    objRect.Y1 = objRect.Y1 + 15
                    strItemName = "仪  器:" & rsItem!仪器 'Nvl(rsItem!检验项目)
                    Call MakeItemTitle(objRect, objFont, strItemName)
                    '产生结果标题
                    mintCount = mintCount + 1
                    Call MakeResultTitle(objRect, objFont, blnMultiline, rsItem)
                    objFont.Bold = False
                    gclsLisReportLib.Font = objFont
    
                    rsResult.Filter = ""
                    '指标及检查结果
                    intRowNum = 0
                    Do While Not rsResult.EOF
                        blnNewPage = False
                        If Nvl(rsResult("显示英文名").Value) <> "" Then
                            strView = "显示"
                            intRowNum = intRowNum + 1
                        Else
                            strView = ""
                        End If
                        strViewRowNum = IIf(Len(strView) <> 0, intRowNum, "")
                        intRows = 1
'                        intColRows(1) = GetLines(Nvl(rsResult("序号").Value), gclsLisReportLib.GetTwipsX(mlngGridColWidth(1)))
                        intColRows(1) = GetLines(Nvl(rsResult("显示英文名").Value), gclsLisReportLib.GetTwipsX(mlngGridColWidth(2)))
                        intColRows(2) = GetLines(Nvl(rsResult("显示中文名").Value), gclsLisReportLib.GetTwipsX(mlngGridColWidth(3)))
                        intColRows(3) = GetLines(Nvl(rsResult("显示结果").Value), gclsLisReportLib.GetTwipsX(mlngGridColWidth(4)))
                        intColRows(4) = GetLines(Nvl(rsResult("显示标志").Value), gclsLisReportLib.GetTwipsX(mlngGridColWidth(4)))
                        intColRows(5) = GetLines(IIf(Nvl(rsResult("临床特征").Value) = "", "", Nvl(rsResult("临床特征").Value) & "  ") & _
                            Nvl(rsResult("结果参考").Value), gclsLisReportLib.GetTwipsX(mlngGridColWidth(4)))
                        intColRows(6) = GetLines(Nvl(rsResult("显示单位").Value), gclsLisReportLib.GetTwipsX(mlngGridColWidth(4)))
    
                        If intColRows(1) < intColRows(2) Then intRows = intColRows(2)
                        If intRows < intColRows(3) Then intRows = intColRows(3)
                        If intRows < intColRows(4) Then intRows = intColRows(4)
                        If intRows < intColRows(5) Then intRows = intColRows(5)
                        If intRows < intColRows(6) Then intRows = intColRows(6)
'                        If intRows < intColRows(7) Then intRows = intColRows(7)
    '                    If Not blnMultiline And objRect.Y1 + 2 * objRect.B0 + intRows * (gclsLisReportLib.GetTextHeight("高") + objRect.R0) - objRect.R0 > gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - gclsLisReportLib.Paper.PageFoot - gclsLisReportLib.Paper.BorderBottom Then
    ''                        intEndRow = intStartRow * 2
    ''                        rsResult.Filter = ""
    ''                        rsResult.Filter = " 序号>" & intStartRow & " And 序号<=" & intEndRow
    ''                        Do While Not rsResult.EOF
    ''                            bln第二列 = True
    ''                            Call MakeResultData(objRect, objFont, rsItem, rsResult, intColRows, intMaxLine, intStartRow, bln第二列, blnMultiline, intItemCount)
    ''                            rsResult.MoveNext
    ''                        Loop
    ''                        intStartRow = intEndRow + 1
    ''                        rsResult.Filter = ""
    ''                        rsResult.Filter = " 序号>=" & intStartRow
    ''                        bln第二列 = False
    '                    End If
                        '计算本页是否够打印本条记录
                        intRows = 1
                        If objRect.Y1 + 2 * objRect.B0 + intRows * (gclsLisReportLib.GetTextHeight("高") + objRect.R0) - objRect.R0 > gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - gclsLisReportLib.Paper.PageFoot Then  '- gclsLisReportLib.Paper.BorderBottom
    
                            '换页，并补加项目标题及结果标题，上页末增加“(续下页)”，补加后在下页第一行显示“(继上页)”
                            
                            If Not blnMultiline Then
                                blnNewPage = True
                                intEndRow = intStartRow * 2
                                rsResult.Filter = ""
                                intMaxLine = intStartRow
                                rsResult.Filter = " 序号>" & intStartRow & " And 序号<=" & intEndRow
                                intStartRow = 1
                                Do While Not rsResult.EOF
                                    
                                    bln第二列 = True
                                    
                                    If intMaxLine Mod 2 = 1 And Not blnAddRow Then
                                        mintCount = mintCount + 1
                                        blnAddRow = True
                                    End If
                                    
                                    Call MakeResultData(objRect, objFont, rsItem, rsResult, intColRows, intMaxLine, intStartRow, bln第二列, blnMultiline, blnNewPage, intItemCount, strViewRowNum)
                                    If Nvl(rsResult("显示英文名").Value) <> "" Then
                                        strView = "显示"
                                        intRowNum = intRowNum + 1
                                    Else
                                        strView = ""
                                    End If
                                    strViewRowNum = IIf(Len(strView) <> 0, intRowNum, "")
                                    rsResult.MoveNext
                                Loop
                                
                                rsResult.Filter = ""
                                rsResult.Filter = " 序号>=" & (intEndRow + 1)
                                
                                intStartRow = 0
                                bln第二列 = False
                                blnNewPage = False
                                mX0 = 0
                                mY0 = 0
                            End If
                            Call SaveRect(objRect, objSvrRect)
                            Call SaveFont(objFont, objSvrFont)
    
                            '补线
                            objFont.ForeColor = USERCOLOR.表格线色
                            objFont.Size = 9
                            gclsLisReportLib.Font = objFont
                            
                            If mSingular <> 0 And mblnSingular Then
                                objRect.Y1 = mSingular
                            End If
                            mSingular = 0
                            objRect.X0 = gclsLisReportLib.Paper.BorderLeft
                            objRect.Y0 = objRect.Y1
                            objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
                            objRect.Y1 = objRect.Y0
                            Call mclsLisReportData.InsertLine(objRect, objFont)
    
                            '补续下页阅读标志
                            objFont.ForeColor = 0
                            objRect.X0 = gclsLisReportLib.Paper.BorderLeft
                            objRect.Y0 = gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.PageFoot - gclsLisReportLib.Paper.BorderBottom '- gclsLisReportLib.GetTwipsY(0.1)
                            objRect.Y1 = 0
                            Call mclsLisReportData.InsertContinueText("(续下页)", objRect, objFont)
    
                            mclsLisReportData.NewPage
                            objRect.Page = objRect.Page + 1
                            objRect.Y0 = gclsLisReportLib.Paper.PageHead + gclsLisReportLib.Paper.BorderTop + gclsLisReportLib.GetTwipsY(0.2)
                            objRect.Y1 = 0
    
                            '补继上页阅读标志
                            objRect.X0 = gclsLisReportLib.Paper.BorderLeft
                            objRect.Y0 = gclsLisReportLib.Paper.BorderTop + gclsLisReportLib.Paper.PageHead + gclsLisReportLib.GetTwipsY(0.1)
                            objRect.Y1 = 0
                            Call mclsLisReportData.InsertContinueText("(继上页)", objRect, objFont)
    
                            '补线
                            objFont.ForeColor = USERCOLOR.表格线色
                            objRect.X0 = gclsLisReportLib.Paper.BorderLeft
                            objRect.Y0 = gclsLisReportLib.Paper.BorderTop + gclsLisReportLib.Paper.PageHead + gclsLisReportLib.GetTwipsY(0.5)
                            objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
                            objRect.Y1 = objRect.Y0
                            Call mclsLisReportData.InsertLine(objRect, objFont)
    
                            objFont.Size = dblFontSize
                            Call SaveRect(objSvrRect, objRect)
                            Call SaveFont(objSvrFont, objFont)
    
                            objRect.Y1 = gclsLisReportLib.Paper.BorderTop + gclsLisReportLib.Paper.PageHead + gclsLisReportLib.GetTwipsY(0.5)
    
                            mintCount = 1
                            If Not ChkRsState(rsResult) Then Call MakeResultTitle(objRect, objFont, blnMultiline, rsItem)
                            blnAddRow = False
    '                        intNewPage = intNewPage - 1
                        End If
    
                        '产生结果数据
                        mintCount = mintCount + 1
                        intStartRow = intStartRow + 1
                        If Not ChkRsState(rsResult) Then
                            mblnSingular = ((rsResult.RecordCount Mod 2) = 1)
                            If blnMultiline Then
                                intMaxLine = rsResult.RecordCount
                            Else
                                intMaxLine = Format(rsResult.RecordCount / 2, "0")
                            End If
                            If intStartRow > intMaxLine And Not blnMultiline Then
                                bln第二列 = True
                                If intMaxLine Mod 2 = 1 And Not blnAddRow Then
                                    mintCount = mintCount + 1
                                    blnAddRow = True
                                End If
                            End If

                            Call MakeResultData(objRect, objFont, rsItem, rsResult, intColRows, intMaxLine, intStartRow, bln第二列, blnMultiline, blnNewPage, intItemCount, strViewRowNum)
        
                            rsResult.MoveNext
                        End If
                    Loop
                End If
            End If
            '补充检验备注
            '补线
            
            objFont.ForeColor = USERCOLOR.表格线色
            objFont.Size = 9
            gclsLisReportLib.Font = objFont
            
            If mSingular <> 0 And mblnSingular Then
                objRect.Y1 = mSingular
            End If
            mSingular = 0
            objRect.X0 = gclsLisReportLib.Paper.BorderLeft
            objRect.Y0 = objRect.Y1
            objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
            
            objRect.Y1 = objRect.Y0
            Call mclsLisReportData.InsertLine(objRect, objFont)
            Set rsAdvice = mclsLisPackage.Get_LisAdvice(Val(Nvl(rsItem!标本id)))
            If Not ChkRsState(rsAdvice) Then
                strItemTmp = Nvl(rsAdvice!评语)
            Else
                strItemTmp = "检验备注:" & Nvl(rsItem!检验备注)
            End If
            If strItemTmp <> "" Then
                objFont.ForeColor = USERCOLOR.黑色
                objFont.Bold = False
                '    objFont.Size = 12
                gclsLisReportLib.Font = objFont
                
                
                objRect.R0 = gclsLisReportLib.GetTwipsX(0.15)
                objRect.X0 = gclsLisReportLib.Paper.BorderLeft
                objRect.Y0 = objRect.Y1 + gclsLisReportLib.GetTwipsX(0.2)
                objRect.Y1 = objRect.Y0 + 2 * objRect.B0 + gclsLisReportLib.GetTextHeight("高") + gclsLisReportLib.GetTwipsX(0.15)
                
                objFont.BackColor = USERCOLOR.项目标题背景色
                objRect.X0 = gclsLisReportLib.Paper.BorderLeft
                objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
                
    '            Call mclsLisReportData.InsertBackColor(objRect, objFont)
                intRows = GetLines(strItemTmp, gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderLeft - gclsLisReportLib.Paper.BorderRight)
                objFont.ForeColor = USERCOLOR.黑色
                objFont.BackColor = USERCOLOR.项目标题背景色
                objRect.Y0 = objRect.Y0 + 30
                objRect.Y1 = objRect.Y0 + 2 * objRect.B0 + intRows * (gclsLisReportLib.GetTextHeight("高") + objRect.R0) - objRect.R0
    '            objRect.Y1 = objRect.Y1
                objFont.WordWarp = True
                strItemTmp = Get_Text(strItemTmp, gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderLeft - gclsLisReportLib.Paper.BorderRight, Abs(Int(-objRect.X0 / gclsLisReportLib.GetTextWidth(Space(1)))))
                Call mclsLisReportData.InsertText(strItemTmp, objRect, objFont)
            End If
            '获取图片
            strImagePath = ""
            If gbln显示图片 Then
                gobjCommFun.ShowFlash "正在读取图片并组织数据，请稍候...."
                strImagePath = mclsDrawLisImage.GetSampleImages(Val(rsItem!标本id), False, 1)
                gobjCommFun.StopFlash
            End If
            '插入图片
            If strImagePath <> "" Then Call MakeLisPicResultData(objRect, objFont, strImagePath)
    '        If strImagePath = "" Then
    '            objFont.ForeColor = USERCOLOR.表格线色
    '            objFont.Size = 9
    '            gclsLisReportLib.Font = objFont
    '
    '            objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    '            objRect.Y0 = objRect.Y1 + gclsLisReportLib.GetTwipsX(0.2)
    '            objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
    '            objRect.Y1 = objRect.Y0
    '            Call mclsLisReportData.InsertLine(objRect, objFont)
    '        End If
            '插入页脚信息
            objFont.Size = 10.5
            objFont.ForeColor = USERCOLOR.黑色
            objFont.Bold = False
            btyType = 2
            '前面图片
            strSign = ""
            strTmp = ""
            
            Set rsSignId = mclsLisPackage.Get_LisSignId(rsItem!审核人)
            If Not ChkRsState(rsSignId) Then
                If mclsLisPackage.IsNewVersion Then
                    strTmp = gobjSystem.ReadLob(100, 15, rsSignId!ID)
                    
                    If strTmp <> "" Then
                        frmSignPic.imgSign.Picture = LoadPicture(strTmp)
                        strSign = gstrSignPath & "\" & rsSignId!ID & ".jpg"
                        SavePicture frmSignPic.imgSign.Picture, strSign
                    End If
                Else
                    Set rsSignPicture = mclsLisPackage.Get_LisSignPicture(rsSignId!ID)
                    
                    strTmp = "zlLisSignPicture" & Format(Now, "yyyymmdd") & Format(100 * Timer, "0") & ".tmp"
                    
                    strTmp = CreateTmpFile(strTmp)
                    strTmp = gobjDatabase.ReadPicture(rsSignPicture, "签名图片", strTmp)
                    frmSignPic.imgSign.Picture = LoadPicture(strTmp)
                    strSign = gstrSignPath & "\" & rsSignId!ID & ".jpg"
                    If frmSignPic.imgSign.Picture <> 0 Then
                        SavePicture frmSignPic.imgSign.Picture, strSign
                    Else
                        strSign = ""
                    End If
                End If
            End If
            Call MakePatientTitle(objRect, objFont, btyType, rsItem, strSign)
            
            '表头空行
            
    '        objFont.ForeColor = USERCOLOR.表格线色
    '        objFont.Size = 9
    '        gclsLisReportLib.Font = objFont
    '
'            objRect.X0 = gclsLisReportLib.Paper.BorderLeft
'            objRect.Y0 = objRect.Y1 + gclsLisReportLib.GetTwipsY(0.5)
'            objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
'            objRect.Y1 = objRect.Y0
    '        Call mclsLisReportData.InsertLine(objRect, objFont)
            
            rsItem.MoveNext
        Loop
        
        objFont.Size = 10.5
        objFont.ForeColor = USERCOLOR.黑色
        objFont.Bold = True
        objFont.ForeColor = USERCOLOR.黑色
        
        objRect.X0 = gclsLisReportLib.Paper.BorderLeft
        objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
        objRect.Y0 = objRect.Y0 + gclsLisReportLib.GetTwipsY(0.7)
        objRect.Y1 = objRect.Y1
        objFont.WordWarp = True
    
        strTmp = "**此结果仅对本标本负责**   您的健康是我们最大的追求，祝您早日康复！  "
        
        Call mclsLisReportData.InsertText(strTmp, objRect, objFont)
        
        objFont.Size = 9
        objFont.ForeColor = USERCOLOR.黑色
        gclsLisReportLib.Font = objFont
        
        '4.保存最后的输出位置，为下一步输出作准备(下面的代码不能删除)
        '------------------------------------------------------------------------------------------------------------------
        objRect.X0 = gclsLisReportLib.Paper.BorderLeft
        objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
        CurrentRect = objRect
    End If
    GetElementData = True
    
    Exit Function
    
    '5.出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SetRect(ByRef objRect As USERRECT, ByVal X0 As Long, ByVal Y0 As Long, ByVal X1 As Long, ByVal Y1 As Long, Optional ByVal B0 As Long, Optional ByVal R0 As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    objRect.X0 = X0
    objRect.Y0 = Y0
    objRect.X1 = X1
    objRect.Y1 = Y1
    objRect.B0 = B0
    objRect.R0 = R0
End Function

Private Function SaveRect(ByRef objFromRect As USERRECT, ByRef objToRect As USERRECT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    objToRect.X0 = objFromRect.X0
    objToRect.Y0 = objFromRect.Y0
    objToRect.X1 = objFromRect.X1
    objToRect.Y1 = objFromRect.Y1
    objToRect.B0 = objFromRect.B0
    objToRect.R0 = objFromRect.R0
    
End Function

Private Function SaveFont(ByRef objFromFont As USERFONT, ByRef objToFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    objToFont.Name = objFromFont.Name
    objToFont.Size = objFromFont.Size
    objToFont.Bold = objFromFont.Bold
    objToFont.Italic = objFromFont.Italic
    objToFont.Underline = objFromFont.Underline
    objToFont.BackColor = objFromFont.BackColor
    objToFont.ForeColor = objFromFont.ForeColor
    
End Function

Private Function MakeItemTitle(ByRef objRect As USERRECT, ByRef objFont As USERFONT, ByVal strItemName As String) As Boolean
    '******************************************************************************************************************
    '功能：产生项目标题
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTmp      As String
    Dim intRows     As Integer
    
    objFont.ForeColor = USERCOLOR.黑色
    objFont.Bold = False
    objFont.Size = 8
    gclsLisReportLib.Font = objFont
    

    objRect.R0 = gclsLisReportLib.GetTwipsX(0.15)
    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.Y0 = objRect.Y1
    objRect.Y1 = objRect.Y0 + 2 * objRect.B0 + gclsLisReportLib.GetTextHeight("高")
    
'    objRect.Y1 = objRect.Y0 + gclsLisReportLib.GetTwipsX(mlngGrIDRowHeight)
'    intRows = GetLines(strItemName, gclsLisReportLib.GetTwipsX(mlngGridColWidth(1) + mlngGridColWidth(2)))
'
'    objRect.Y1 = objRect.Y0 + 2 * objRect.B0 + intRows * (gclsLisReportLib.GetTextHeight("高") + objRect.R0) - objRect.R0
    
    objFont.BackColor = USERCOLOR.表格背景色
    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
    
    Call mclsLisReportData.InsertBackColor(objRect, objFont)

    
'    objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1) + mlngGridColWidth(2))
'    strItemName = Get_Text(strItemName, gclsLisReportLib.GetTwipsX(mlngGridColWidth(1) + mlngGridColWidth(2)), Abs(Int(-objRect.X0 / gclsLisReportLib.GetTextWidth(Space(1)))))
    Call mclsLisReportData.InsertText(strItemName, objRect, objFont)
    
    gclsLisReportLib.Font = objFont
    
End Function

Private Function MakeResultTitle(ByRef objRect As USERRECT, ByRef objFont As USERFONT, ByVal blnMultiline As Boolean, ByVal rsItem As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：产生结果标题
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTmp          As String
    Dim objSvrRect      As USERRECT
    Dim intMultileCol   As Integer
    
    On Error GoTo errHand
    
    objFont.Bold = False
    objFont.BackColor = USERCOLOR.表格背景色
    objFont.ForeColor = USERCOLOR.黑色
    gclsLisReportLib.Font = objFont
    
    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.Y0 = objRect.Y1
    objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
'    objRect.Y1 = objRect.Y0 + gclsLisReportLib.GetTwipsX(mlngGrIDRowHeight)
    objRect.Y1 = objRect.Y0 + 2 * objRect.B0 + gclsLisReportLib.GetTextHeight("高")

    If mintCount Mod 2 = 0 Then
        objFont.BackColor = USERCOLOR.表格背景色
        objRect.X0 = gclsLisReportLib.Paper.BorderLeft
        objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
        Call mclsLisReportData.InsertBackColor(objRect, objFont)
    End If
        
'    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
'    objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
'    Call mclsLisReportData.InsertText("指标名称", objRect, objFont)
    
    If rsItem("类别").Value = "检查" Then
        
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(2) + mlngGridColWidth(3) + mlngGridColWidth(4) + mlngGridColWidth(5))
        Call mclsLisReportData.InsertText("检查结果", objRect, objFont)
        
        objFont.ForeColor = USERCOLOR.表格线色
        objRect.X0 = gclsLisReportLib.Paper.BorderLeft + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
        objRect.X1 = objRect.X0
        Call mclsLisReportData.InsertLine(objRect, objFont)
    Else
        If blnMultiline Then
            objRect.X0 = gclsLisReportLib.Paper.BorderLeft
'            objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
'            Call mclsLisReportData.InsertText("", objRect, objFont)
            
'            objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
            objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
            Call mclsLisReportData.InsertText("英文名", objRect, objFont)
'
            objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
            objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(2))
            Call mclsLisReportData.InsertText("项目名称", objRect, objFont)
'
            objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(2))
            objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(3))
            Call mclsLisReportData.InsertText("结   果", objRect, objFont)

            objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(3))
            objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(4))
            Call mclsLisReportData.InsertText(IIf(mblnEnzymeMarker, "OD值", ""), objRect, objFont)

            objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(4))
            objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(5))
            Call mclsLisReportData.InsertText(IIf(mblnEnzymeMarker, "cut off", "参 考 值"), objRect, objFont)

            objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(5))
            objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(6))
            Call mclsLisReportData.InsertText(IIf(mblnEnzymeMarker, "S/C.0", "单位"), objRect, objFont)
            
            objFont.ForeColor = USERCOLOR.表格线色
            objRect.X0 = gclsLisReportLib.Paper.BorderLeft + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
            objRect.X1 = objRect.X0
'            Call mclsLisReportData.InsertLine(objRect, objFont)
'
            objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(2))
            objRect.X1 = objRect.X0
'            Call mclsLisReportData.InsertLine(objRect, objFont)
'
            objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(3))
            objRect.X1 = objRect.X0
'            Call mclsLisReportData.InsertLine(objRect, objFont)

            objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(4))
            objRect.X1 = objRect.X0
'            Call mclsLisReportData.InsertLine(objRect, objFont)

            objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(5))
            objRect.X1 = objRect.X0
'            Call mclsLisReportData.InsertLine(objRect, objFont)

            objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(6))
            objRect.X1 = objRect.X0
'            Call mclsLisReportData.InsertLine(objRect, objFont)
        
        Else
            For intMultileCol = 0 To 1
                If intMultileCol = 0 Then
                    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
                Else
                    objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(6)) + 100
                End If
'                objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
'                Call mclsLisReportData.InsertText("", objRect, objFont)
'
'                objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
                objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
                Call mclsLisReportData.InsertText("英文名", objRect, objFont)
                '
                objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
                objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(2))
                Call mclsLisReportData.InsertText("项目名称", objRect, objFont)
                '
                objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(2))
                objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(3))
                Call mclsLisReportData.InsertText("结   果", objRect, objFont)
                
                objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(3))
                objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(4))
'                Call mclsLisReportData.InsertText("", objRect, objFont)
                
                objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(4))
                objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(5))
                Call mclsLisReportData.InsertText("参 考 值", objRect, objFont)
                
                objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(5))
                objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(6))
                Call mclsLisReportData.InsertText("单位", objRect, objFont)
            Next
            For intMultileCol = 0 To 1
                objFont.ForeColor = USERCOLOR.表格线色
                If intMultileCol = 0 Then
                    objRect.X0 = gclsLisReportLib.Paper.BorderLeft + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
                Else
                    objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
                End If
                objRect.X1 = objRect.X0
'                Call mclsLisReportData.InsertLine(objRect, objFont)

                objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(2))
                objRect.X1 = objRect.X0
'                Call mclsLisReportData.InsertLine(objRect, objFont)

                objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(3))
                objRect.X1 = objRect.X0
'                Call mclsLisReportData.InsertLine(objRect, objFont)

                objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(4))
                objRect.X1 = objRect.X0
'                Call mclsLisReportData.InsertLine(objRect, objFont)

                objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(5))
                objRect.X1 = objRect.X0
'                Call mclsLisReportData.InsertLine(objRect, objFont)

'                objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(6))
'                objRect.X1 = objRect.X0
'                Call mclsLisReportData.InsertLine(objRect, objFont)

                If intMultileCol = 0 Then
                    objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(6))
                    objRect.X1 = objRect.X0
                    Call mclsLisReportData.InsertLine(objRect, objFont)
                End If
            Next
        End If
    End If
    
    MakeResultTitle = True
    
    Exit Function
    
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function MakeResultData(ByRef objRect As USERRECT, ByRef objFont As USERFONT, _
                            ByVal rsItem As ADODB.Recordset, ByVal rsResult As ADODB.Recordset, Optional ByVal intColRows As Variant, _
                            Optional ByVal intMaxLine As Integer, Optional ByVal intCurrLine As Integer, Optional ByVal bln第二列 As Boolean, _
                            Optional ByVal blnMultiline As Boolean, Optional ByVal blnNewPage As Boolean, Optional ByVal intItemCount As Integer, _
                            Optional ByVal strRowNum As String) As Boolean
    '******************************************************************************************************************
    '功能：产生结果数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTmp      As String
    Dim objSvrRect  As USERRECT
    Dim intRows     As Integer
    Dim strSign     As String
    Dim sngOldY0    As Single
    Dim sngOldY1    As Single
    
    On Error GoTo errHand
    
    intRows = intColRows(1)
    
    '填写内容之前，先进行最大行高的计算
    gclsLisReportLib.Font = objFont
    
    objFont.Bold = False
    objFont.HAlignment = 1
    
    objRect.B0 = gclsLisReportLib.GetTwipsX(0.1)
    objRect.R0 = gclsLisReportLib.GetTwipsX(0.15)
    

    If intColRows(1) < intColRows(2) Then intRows = intColRows(2)
    If intRows < intColRows(3) Then intRows = intColRows(3)
    If intRows < intColRows(4) Then intRows = intColRows(4)
    If intRows < intColRows(5) Then intRows = intColRows(5)
    If intRows < intColRows(6) Then intRows = intColRows(6)
'    If intRows < intColRows(7) Then intRows = intColRows(7)
    
    If rsItem("类别").Value = "检查" Then
        
        strTmp = Nvl(rsResult("结果").Value)
        Dim varData As Variant
        varData = Split(strTmp, vbCrLf)
        If UBound(varData) = 1 Then strTmp = Replace(strTmp, vbCrLf, "") '只有一行回车,就去掉ext(strTmp, objRect, objFont)
'    Call AppendPrintData(rsData, "项目", "文本", objRect, objFont, gclsLisReportLib.Paper, , True, Nvl(rs("指标").Value), , , intColRows(1))

        strTmp = Replace(strTmp, Chr(10) & Chr(13), "[peispeispeispeis]")
        strTmp = Replace(strTmp, Chr(13) & Chr(10), "[peispeispeispeis]")
        strTmp = Replace(strTmp, Chr(10), "[peispeispeispeis]")
        strTmp = Replace(strTmp, Chr(13), "[peispeispeispeis]")
        strTmp = Replace(strTmp, "[peispeispeispeis]", vbCrLf)
        strTmp = Replace(strTmp, vbTab, Space(2))
        If Nvl(rsResult("单位").Value) <> "" Then strTmp = strTmp & " " & Nvl(rsResult("单位").Value)
        
        
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1)) + objRect.R0
        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(2) + mlngGridColWidth(3) + mlngGridColWidth(4) + mlngGridColWidth(5))
'        Call AppendPrintData(rsData, "项目", "文本", objRect, objFont, gclsLisReportLib.Paper, 1, True, strTmp, , , intColRows(2))
        If intRows > 1 Then
            objFont.MuliLine = True
            sngOldY0 = objRect.Y0
            sngOldY1 = objRect.Y1
        End If
        Call mclsLisReportData.InsertText(strTmp, objRect, objFont)
        If intRows > 1 Then
            objFont.MuliLine = False
            objRect.Y0 = sngOldY0
            objRect.Y1 = sngOldY1
        End If
    Else
'        Debug.Print Nvl(rsResult("序号").Value)
        intRows = 1
        If intCurrLine <= intMaxLine And Not blnNewPage Then
            objRect.X0 = gclsLisReportLib.Paper.BorderLeft
            objRect.Y0 = objRect.Y1
            If intItemCount > 1 And mY0 = 0 Then
'                If mblnSingular Then mSingular = objRect.Y1
                mY0 = objRect.Y0
''            Else
''                objRect.Y0 = objRect.Y1
            End If
        Else
            If bln第二列 And mX0 = 0 Then
                '第二列开始的检验结果坐标
                objRect.X0 = objRect.X0 + 100 '+ gclsLisReportLib.GetTwipsX(mlngGridColWidth(7))
                If intItemCount > 1 Then
                    If mblnSingular Then mSingular = objRect.Y1
                    objRect.Y0 = mY0
                Else
                    If mblnSingular Then mSingular = objRect.Y1
                    objRect.Y0 = 2313
                End If
                mX0 = objRect.X0
            Else
                objRect.X0 = mX0
                objRect.Y0 = objRect.Y1
            End If
        End If
        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
        objRect.Y1 = objRect.Y0 + 2 * objRect.B0 + gclsLisReportLib.GetTextHeight("高")
        objRect.Y1 = objRect.Y0 + 2 * objRect.B0 + intRows * (gclsLisReportLib.GetTextHeight("高") + objRect.R0) - objRect.R0
'        If bln第二列 Then
'            If (mintCount - 1) Mod 2 = 0 Then
'                objFont.BackColor = USERCOLOR.表格背景色
'                objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
'                Call mclsLisReportData.InsertBackColor(objRect, objFont)
'            End If
'        Else
        If mintCount Mod 2 = 0 Then
            objFont.BackColor = USERCOLOR.表格背景色
            objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
            Call mclsLisReportData.InsertBackColor(objRect, objFont)
        End If
'        End If
        '数据
        '------------------------------------------------------------------------------------------------------------------
        objFont.ForeColor = USERCOLOR.黑色
    
        objFont.Rows = intColRows(1)
    
'            strTmp = Get_Text(Nvl(rsResult("序号").Value), gclsLisReportLib.GetTwipsX(mlngGridColWidth(3)), Abs(Int(-objRect.X0 / gclsLisReportLib.GetTextWidth(Space(1)))))
        Call mclsLisReportData.InsertText(Mid(Nvl(rsResult("显示英文名").Value), 1, 5), objRect, objFont)
    
'        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1)) + objRect.R0
'        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(2))
'
'        objFont.Rows = 1
''            objFont.WordWarp = True
''            strTmp = Get_Text(Nvl(rsResult("显示英文名").Value), gclsLisReportLib.GetTwipsX(mlngGridColWidth(2)), Abs(Int(-objRect.X0 / gclsLisReportLib.GetTextWidth(Space(1)))))
'
'        Call mclsLisReportData.InsertText(Mid(Nvl(rsResult("显示英文名").Value), 1, 5), objRect, objFont)
'
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(2))
        objFont.Rows = 1
'            objFont.WordWarp = True
'            strTmp = Get_Text(Nvl(rsResult("显示中文名").Value), gclsLisReportLib.GetTwipsX(mlngGridColWidth(3)), Abs(Int(-objRect.X0 / gclsLisReportLib.GetTextWidth(Space(1)))))
        If Not blnMultiline Then
            strTmp = Mid(Nvl(rsResult("显示中文名").Value), 1, 10)
        Else
            strTmp = Nvl(rsResult("显示中文名").Value)
        End If
        Call mclsLisReportData.InsertText(strTmp, objRect, objFont)
        
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(2))
        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(3))
        objFont.Rows = 1
'            objFont.WordWarp = True
'            strTmp = Get_Text(Nvl(rsResult("显示结果").Value), gclsLisReportLib.GetTwipsX(mlngGridColWidth(4)), Abs(Int(-objRect.X0 / gclsLisReportLib.GetTextWidth(Space(1)))))
        Call mclsLisReportData.InsertText(Nvl(rsResult("显示结果").Value), objRect, objFont)
        
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(3))
        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(4))
        objFont.Rows = 1
'            objFont.WordWarp = True
        
'            strTmp = Get_Text(Nvl(rsResult("显示标志").Value), gclsLisReportLib.GetTwipsX(mlngGridColWidth(5)), Abs(Int(-objRect.X0 / gclsLisReportLib.GetTextWidth(Space(1)))))
        If mblnEnzymeMarker Then
            Call mclsLisReportData.InsertText(Nvl(rsResult("显示OD").Value), objRect, objFont)
        Else
            objFont.ForeColor = USERCOLOR.红色
            Call mclsLisReportData.InsertText(Nvl(rsResult("显示标志").Value), objRect, objFont)
        End If
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(4))
        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(5))
        objFont.Rows = 1
'            objFont.WordWarp = True
        objFont.ForeColor = USERCOLOR.黑色
'            strTmp = Get_Text(IIf(Nvl(rsResult("临床特征").Value) = "", "", Nvl(rsResult("临床特征").Value) & " ") & Nvl(rsResult("结果参考").Value), gclsLisReportLib.GetTwipsX(mlngGridColWidth(6)), Abs(Int(-objRect.X0 / gclsLisReportLib.GetTextWidth(Space(1)))))
        If mblnEnzymeMarker Then
            Call mclsLisReportData.InsertText(Nvl(rsResult("显示CUTOFF").Value), objRect, objFont)
        Else
            Call mclsLisReportData.InsertText(IIf(Nvl(rsResult("临床特征").Value) = "", "", Nvl(rsResult("临床特征").Value) & " ") & Nvl(rsResult("结果参考").Value), objRect, objFont)
        End If
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(5))
        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(6))
        objFont.Rows = 1
'        objFont.WordWarp = True
'            strTmp = Get_Text(Nvl(rsResult("显示单位").Value), gclsLisReportLib.GetTwipsX(mlngGridColWidth(7)), Abs(Int(-objRect.X0 / gclsLisReportLib.GetTextWidth(Space(1)))))
        If mblnEnzymeMarker Then
            Call mclsLisReportData.InsertText(Nvl(rsResult("显示SCO").Value), objRect, objFont)
        Else
            Call mclsLisReportData.InsertText(Nvl(rsResult("显示单位").Value), objRect, objFont)
        End If
        objFont.Rows = 1
        objFont.WordWarp = False
        objFont.HAlignment = 1
    End If

    
    '画纵线
    '------------------------------------------------------------------------------------------------------------------

    objFont.ForeColor = USERCOLOR.表格线色
    
    If rsItem("类别").Value = "检查" Then
        objRect.X0 = gclsLisReportLib.Paper.BorderLeft + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
        objRect.X1 = objRect.X0
        Call mclsLisReportData.InsertLine(objRect, objFont)
    Else
        If intCurrLine <= intMaxLine And Not blnNewPage Then
            objRect.X0 = gclsLisReportLib.Paper.BorderLeft + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
        Else
            objRect.X0 = mX0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1))
        End If
        objRect.X1 = objRect.X0
'        Call mclsLisReportData.InsertLine(objRect, objFont)
    
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(2))
        objRect.X1 = objRect.X0
'        Call mclsLisReportData.InsertLine(objRect, objFont)
        
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(3))
        objRect.X1 = objRect.X0
'        Call mclsLisReportData.InsertLine(objRect, objFont)
        
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(4))
        objRect.X1 = objRect.X0
'        Call mclsLisReportData.InsertLine(objRect, objFont)
'
'        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(5))
'        objRect.X1 = objRect.X0
'    '        Call mclsLisReportData.InsertLine(objRect, objFont)

        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(5))
        objRect.X1 = objRect.X0
''        Call mclsLisReportData.InsertLine(objRect, objFont)
        If Not bln第二列 And Not blnMultiline Then
            objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(6))
            objRect.X1 = objRect.X0
            Call mclsLisReportData.InsertLine(objRect, objFont)
        End If
    End If
    
    MakeResultData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Private Function MakeBacilliResultData(ByRef objRect As USERRECT, ByRef objFont As USERFONT, ByVal rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim objclsGrID  As clsGrid
    Dim objSvrRect  As USERRECT
    Dim objSvrFont  As USERFONT
    Dim strTmp      As String
    Dim intColRows  As Variant
    
    On Error GoTo errHand
    
    intColRows = Split("1;1;1;1", ";")
    
    '获取当前记录的输出行数
    intColRows(0) = Get_MaxRows(Nvl(rs("细菌名称").Value), 18.5)
    If Trim(rs("细菌结果").Value) = "" Then
        intColRows(1) = 0
    Else
        intColRows(1) = Get_MaxRows(Nvl(rs("细菌结果").Value), 16.5)
    End If
    intColRows(2) = Get_MaxRows(Nvl(rs("耐药机制").Value), 16.5)
    intColRows(3) = Get_MaxRows(Nvl(rs("培养描述").Value), 16.5)

    
    '计算长度进行换页
    If (objRect.Y1 + 2 * objRect.B0 + (Val(intColRows(0)) + Val(intColRows(1)) + Val(intColRows(2)) + Val(intColRows(3))) * (gclsLisReportLib.GetTwipsY("0.6") + objRect.R0) - objRect.R0) + gclsLisReportLib.GetTwipsY(0.4) > (gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - gclsLisReportLib.Paper.BorderTop) Then
        
        '超过了本页高度
        Call ContinuePage(objRect, objFont, False)
        
        objRect.Y0 = gclsLisReportLib.Paper.BorderTop + gclsLisReportLib.Paper.PageHead + gclsLisReportLib.GetTwipsY(0.5)
        objRect.Y1 = objRect.Y0
    End If
    
    '产生细菌结果
    objFont.ForeColor = USERCOLOR.表格线色
    '第1行
    Set objclsGrID = New clsGrid
    Call objclsGrID.InitGrID(1, USERCOLOR.黑色, USERCOLOR.白色, objFont.Name, objFont.Size, "000000")
    Call objclsGrID.SetGrIDRowFormat(0.6 * intColRows(0))
    Call objclsGrID.SetGrIDColFormat("18")
    
    strTmp = rs.AbsolutePosition & "." & Nvl(rs("细菌名称").Value)
    Call objclsGrID.SetGrIDRowData(1, strTmp)
    objFont.Rows = intColRows(0)
    objFont.Bold = True
    objclsGrID.Cell(WordWarp, 1, 1) = 1
    
    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.Y0 = objRect.Y1

    Call mclsLisReportData.InsertGrID(objclsGrID, objRect, objFont)
    objFont.Bold = False
    If Trim(rs("细菌结果").Value) <> "" Then
        '第2行
        Set objclsGrID = New clsGrid
        Call objclsGrID.InitGrID(1, USERCOLOR.黑色, USERCOLOR.白色, objFont.Name, objFont.Size, "100000")
        Call objclsGrID.SetGrIDRowFormat(0.6 * intColRows(1))
        Call objclsGrID.SetGrIDColFormat("2.5;17.5")
        
        strTmp = "结   果:;" & Nvl(rs("细菌结果").Value)
        Call objclsGrID.SetGrIDRowData(1, strTmp)
        objFont.Rows = intColRows(1)
        objclsGrID.Cell(WordWarp, 1, 1) = 1
        objclsGrID.Cell(WordWarp, 1, 2) = 1
    
        objRect.X0 = gclsLisReportLib.Paper.BorderLeft
        objRect.Y0 = objRect.Y1
    
        Call mclsLisReportData.InsertGrID(objclsGrID, objRect, objFont)
    End If
    '第3行
    Set objclsGrID = New clsGrid
    If Trim(rs("细菌结果")) <> "" Then
        Call objclsGrID.InitGrID(1, USERCOLOR.黑色, USERCOLOR.白色, objFont.Name, objFont.Size, "000000")
    Else
        Call objclsGrID.InitGrID(1, USERCOLOR.黑色, USERCOLOR.白色, objFont.Name, objFont.Size, "100000")
    End If
    Call objclsGrID.SetGrIDRowFormat(0.6 * intColRows(2))
    Call objclsGrID.SetGrIDColFormat("2.5;17.5")
    
    strTmp = "耐药机制:;" & Nvl(rs("耐药机制").Value)
    Call objclsGrID.SetGrIDRowData(1, strTmp)
    objFont.Rows = intColRows(2)
    objclsGrID.Cell(WordWarp, 1, 1) = 1
    objclsGrID.Cell(WordWarp, 1, 2) = 1

    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.Y0 = objRect.Y1

    Call mclsLisReportData.InsertGrID(objclsGrID, objRect, objFont)
        
    '第4行
    Set objclsGrID = New clsGrid
    Call objclsGrID.InitGrID(1, USERCOLOR.黑色, USERCOLOR.白色, objFont.Name, objFont.Size, "000000")
    Call objclsGrID.SetGrIDRowFormat(0.6 * intColRows(3))
    Call objclsGrID.SetGrIDColFormat("2.5;17.5")
    
    strTmp = "培养描述:;" & Nvl(rs("培养描述").Value)
    Call objclsGrID.SetGrIDRowData(1, strTmp)
    objFont.Rows = intColRows(3)
    objclsGrID.Cell(WordWarp, 1, 1) = 1
    objclsGrID.Cell(WordWarp, 1, 2) = 1

    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.Y0 = objRect.Y1

    Call mclsLisReportData.InsertGrID(objclsGrID, objRect, objFont)
    
    objFont.Rows = 1
    
    MakeBacilliResultData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function MakeAntibioticResultTitle(ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '功能：产生抗生素结果标题
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objclsGrID As clsGrid
    Dim intColRows As Variant

    On Error GoTo errHand

    intColRows = Split("1;1;1;1;1;1", ";")

    objRect.B0 = gclsLisReportLib.GetTwipsX(0.1)
    objRect.R0 = gclsLisReportLib.GetTwipsX(0.15)
    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.Y0 = objRect.Y1
    objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
    objRect.Y1 = 0

    '设置起始位置
    If (objRect.Y0 + gclsLisReportLib.GetTwipsY(1.2)) > (gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - gclsLisReportLib.Paper.PageFoot) Then
        mclsLisReportData.NewPage
        objRect.Page = objRect.Page + 1
        objRect.Y0 = gclsLisReportLib.Paper.BorderTop + gclsLisReportLib.Paper.PageHead
    End If
    
    objFont.ForeColor = USERCOLOR.表格线色
    Set objclsGrID = New clsGrid
    Call objclsGrID.InitGrID(1, USERCOLOR.黑色, USERCOLOR.白色, objFont.Name, objFont.Size, "110000")
    Call objclsGrID.SetGrIDRowFormat("0.6")
    Call objclsGrID.SetGrIDColFormat("6.5;1;1.8;0.2")
'    Call objclsGrID.SetGrIDColFormat("5.5;2.5;2")
    Call objclsGrID.SetGrIDRowData(1, "抗生素名;结果;结果类型;")

    objFont.Bold = True
    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    Call mclsLisReportData.InsertGrID(objclsGrID, objRect, objFont)
    objFont.Bold = False

    Set objclsGrID = New clsGrid
    Call objclsGrID.InitGrID(1, USERCOLOR.黑色, USERCOLOR.白色, objFont.Name, objFont.Size, "100000")
    Call objclsGrID.SetGrIDRowFormat("0.6")
    Call objclsGrID.SetGrIDColFormat("6.5;1;3.1;0.2")
'    Call objclsGrID.SetGrIDColFormat("6;2.5;2")
    Call objclsGrID.SetGrIDRowData(1, "  抗生素名;结果;结果类型;")

    objRect.X0 = gclsLisReportLib.Paper.BorderLeft + gclsLisReportLib.GetTwipsX(9.2)
    objFont.Bold = True
    Call mclsLisReportData.InsertGrID(objclsGrID, objRect, objFont)
    objFont.Bold = False
    objFont.ForeColor = USERCOLOR.黑色

    MakeAntibioticResultTitle = True

    Exit Function

errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function MakeAntibioticResultData(ByRef objRect As USERRECT, ByRef objFont As USERFONT, ByVal rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：输出抗生素结果
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim objSvrRect  As USERRECT
    Dim objSvrFont  As USERFONT
    Dim objclsGrID  As clsGrid
    Dim strTmp      As String
    Dim intColRows  As Variant
    Dim lngRows     As Long
    Dim strCol      As String
    Dim strFlag     As String
    
    On Error GoTo errHand
    
    '填写内容之前，先进行最大行高的计算
    gclsLisReportLib.Font = objFont
    intColRows = Split("1;1;1;1;1;1", ";")
    
    objFont.VAlignment = 2
    objRect.B0 = gclsLisReportLib.GetTwipsX(0.1)
    objRect.R0 = gclsLisReportLib.GetTwipsX(0.15)
    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.Y0 = objRect.Y1
        
    Do While Not rs.EOF
'        6.5;1;2
        intColRows(0) = Get_MaxRows(Nvl(rs("抗生素1").Value), 6.5)
        intColRows(1) = Get_MaxRows(Nvl(rs("结果1").Value), 1)
        intColRows(2) = Get_MaxRows(Nvl(rs("耐药性1").Value), 2)
        intColRows(3) = Get_MaxRows(Nvl(rs("抗生素2").Value), 6.5)
        intColRows(4) = Get_MaxRows(Nvl(rs("结果2").Value), 1)
        intColRows(5) = Get_MaxRows(Nvl(rs("耐药性2").Value), 2)
        
        lngRows = intColRows(0)
        
        If lngRows < intColRows(1) Then lngRows = intColRows(1)
        If lngRows < intColRows(2) Then lngRows = intColRows(2)
        If lngRows < intColRows(3) Then lngRows = intColRows(3)
        If lngRows < intColRows(4) Then lngRows = intColRows(4)
        If lngRows < intColRows(5) Then lngRows = intColRows(5)
        lngRows = 1
        '计算长度进行换页
        If (objRect.Y1 + 2 * objRect.B0 + lngRows * (gclsLisReportLib.GetTextHeight("高") + objRect.R0) - objRect.R0) + gclsLisReportLib.GetTwipsY(0.4) > (gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - gclsLisReportLib.Paper.BorderTop) Then

            '超过了本页高度
            Call ContinuePage(objRect, objFont, False)
            
            objRect.Y0 = gclsLisReportLib.Paper.BorderTop + gclsLisReportLib.Paper.PageHead + gclsLisReportLib.GetTwipsY(0.5)
            objRect.Y1 = objRect.Y0

            Call MakeAntibioticResultTitle(objRect, objFont)
        End If
        objFont.VAlignment = 1
        objFont.HAlignment = 1
        
        objRect.Y0 = objRect.Y1
        objRect.Y1 = objRect.Y0 + 2 * objRect.B0 + lngRows * (gclsLisReportLib.GetTextHeight("高") + objRect.R0) - objRect.R0
        
        If rs.AbsolutePosition Mod 2 = 1 Then
            objFont.BackColor = USERCOLOR.表格背景色
            objRect.X0 = gclsLisReportLib.Paper.BorderLeft
            objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
            Call mclsLisReportData.InsertBackColor(objRect, objFont)
        End If
        
        objFont.ForeColor = USERCOLOR.黑色
        objRect.X0 = gclsLisReportLib.Paper.BorderLeft
        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsY(6.5)
        
'        objFont.WordWarp = True
        objFont.Rows = 1 ' intColRows(0)
        strTmp = Mid(Nvl(rs("抗生素1").Value), 1, 20)
        Call mclsLisReportData.InsertText(strTmp, objRect, objFont)
        
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsY(6.5)
        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsY(1)
        
'        objFont.WordWarp = True
        objFont.Rows = 1 'intColRows(1)
        strTmp = Nvl(rs("结果1").Value)
        Call mclsLisReportData.InsertText(strTmp, objRect, objFont)
        
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsY(1)
        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsY(1.8)
        
'        objFont.WordWarp = True
        objFont.Rows = 1 ' intColRows(2)
        strTmp = Nvl(rs("耐药性1").Value)

        Call mclsLisReportData.InsertText(strTmp, objRect, objFont)
        
        If InStr(strTmp, "敏感") > 0 Then
            strFlag = "*"
            objFont.ForeColor = USERCOLOR.红色
        Else
            strFlag = ""
            objFont.ForeColor = USERCOLOR.黑色
        End If
        
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsY(1.1)
        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsY(0.2)
        Call mclsLisReportData.InsertText(strFlag, objRect, objFont)
        
        objFont.ForeColor = USERCOLOR.黑色
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsY(0.9)
        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsY(6.5)
                
'        objFont.WordWarp = True
        objFont.Rows = 1 'intColRows(3)
        strTmp = " " & Mid(Nvl(rs("抗生素2").Value), 1, 30)
        Call mclsLisReportData.InsertText(strTmp, objRect, objFont)
        
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsY(6.5)
        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsY(1)
        
'        objFont.WordWarp = True
        objFont.Rows = 1 ' intColRows(4)
        strTmp = Nvl(rs("结果2").Value)
        Call mclsLisReportData.InsertText(strTmp, objRect, objFont)
        
        
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsY(1)
        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsY(3.1)
        
'        objFont.WordWarp = True
        objFont.Rows = 1 'intColRows(5)
        strTmp = Nvl(rs("耐药性2").Value)

        Call mclsLisReportData.InsertText(strTmp, objRect, objFont)
        
        If InStr(strTmp, "敏感") > 0 Then
            strFlag = "*"
            objFont.ForeColor = USERCOLOR.红色
        Else
            strFlag = ""
            objFont.ForeColor = USERCOLOR.黑色
        End If
        
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTwipsY(1.1)
        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsY(0.2)
        Call mclsLisReportData.InsertText(strFlag, objRect, objFont)
        objFont.ForeColor = USERCOLOR.黑色
        '画纵线
        '------------------------------------------------------------------------------------------------------------------
        objRect.X0 = gclsLisReportLib.Paper.BorderLeft + gclsLisReportLib.GetTwipsY(6 + 2 + 1.5)
        objRect.X1 = objRect.X0
        objRect.Y1 = objRect.Y1 '+ gclsLisReportLib.GetTextHeight("高") '+ objRect.R0 ' + gclsLisReportLib.GetTwipsY(0.1)
        objFont.ForeColor = USERCOLOR.表格线色
        Call mclsLisReportData.InsertLine(objRect, objFont)
        objFont.ForeColor = USERCOLOR.黑色
        objFont.LineStyle = 0
        
        rs.MoveNext
    Loop
    '行结束线
    '------------------------------------------------------------------------------------------------------------------
    objRect.Y0 = objRect.Y1
    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
'    objFont.LineStyle = 1
    objFont.ForeColor = USERCOLOR.表格线色
    Call mclsLisReportData.InsertLine(objRect, objFont)
    objFont.ForeColor = USERCOLOR.黑色
    objFont.LineStyle = 0
    objFont.ForeColor = USERCOLOR.黑色
    
    objFont.WordWarp = False
    objFont.Rows = 1
    MakeAntibioticResultData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Get_MaxRows(ByVal strValue As String, ByVal strCol As String) As Integer

    Dim varValue As Variant
    Dim varCol As Variant
    Dim intIndex As Integer
    Dim intRows As Integer
    
    intRows = 1
    On Error GoTo errHand
    varValue = Split(strValue, ";")
    varCol = Split(strCol, ";")
    
    If UBound(varValue) = UBound(varCol) Then
        For intIndex = 0 To UBound(varValue)
            If intRows < GetLines(varValue(intIndex), gclsLisReportLib.GetTwipsY(varCol(intIndex))) Then
                intRows = GetLines(varValue(intIndex), gclsLisReportLib.GetTwipsY(varCol(intIndex)))
            End If
        Next
    End If
    Get_MaxRows = intRows
    Exit Function
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetLines(ByVal strText As String, ByVal lngCX As Long) As Integer
    '******************************************************************************************************************
    '功能：获取需要的行数，因为有可能要换行
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngChars As Long
    Dim lngTextLength As Long
    Dim lngChar As Long
    Dim blnCrlf As Boolean
    Dim strChar As String
    Dim strLastChar As String
    Dim strNextChar As String
    Dim strLineText As String
    Dim lngRow As Long
    
    lngRow = 1
    lngTextLength = Len(strText)
    
    For lngChar = 1 To lngTextLength
        
        blnCrlf = False
        
        strChar = Mid(strText, lngChar, 1)
        strLastChar = strChar
        
        Select Case Asc(strChar)
        Case 13
            '需要再判断下一个字符是否为换行符
            If lngChar + 1 <= lngTextLength Then
                strNextChar = Mid(strText, lngChar + 1, 1)
                        
                If Asc(strNextChar) = 10 Then
                    '是回车换行符
                    
                    strLastChar = vbCrLf
                    
                    '判断此回车换行符所处的位置，主要判断它的前一字符是不是也是回车换行符或首字符，如果是的则为新行
                    If lngChar = 1 Or strLastChar = vbCrLf Then
                        blnCrlf = True
                    End If
                    
                    '循环变量累加1个计数器
                    lngChar = lngChar + 1
                End If
            
            End If
        Case 10
            '直接换行
            blnCrlf = True
        End Select
        
        If gclsLisReportLib.GetTextWidth(strLineText & strChar) > lngCX Or blnCrlf Then
            
            lngRow = lngRow + 1
            If blnCrlf = False Then
                strLineText = strChar
            Else
                strLineText = ""
            End If
        Else
            strLineText = strLineText & strChar
        End If
    Next
    
    GetLines = lngRow
End Function

Private Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, ByVal RegionHeight As Long, Rows As Integer, Cols As Integer)
    '******************************************************************************************************************
    '功能：根据输入的图像数量，图像区域的宽度和高度，返回最佳的图像排列行数和列数
    '参数： ImageCount－－图像数量
    '       RegionWidth--图像显示区域的宽度
    '       RegionHeight--图像显示区域的高度
    '       Rows－－[返回]最佳行数
    '       Cols－－[返回]最佳列数
    '返回：返回最佳行数Rows，最佳列数Cols
    '******************************************************************************************************************
    Dim iCols As Integer, iRows As Integer

    On Error GoTo err
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))

    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    Do While iRows * iCols > ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols - 1
        Else
            iRows = iRows - 1
        End If
    Loop

    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    Rows = iRows: Cols = iCols
err:
End Sub

Public Function GetPicturePrintHeight(ByVal strFile As String, ByVal cx As Long, ByVal cy As Long) As Single
    '******************************************************************************************************************
    '功能：根据相册大小自动等比例缩放相片文件
    '参数：缩放前的相片文件
    '返回：缩放后的相片文件
    '******************************************************************************************************************
    Dim strTmp As String
    Dim objMap As StdPicture
    Dim W As Single
    Dim H As Single
    Dim sglPerW As Single
    Dim sglPerH As Single
    Dim sglPer As Single
    Dim X As Long
    Dim Y As Long
    
    On Error GoTo errHand
    
    If strFile = "" Then Exit Function
    
    Set objMap = VB.LoadPicture(strFile)
    
    W = objMap.Width * 0.566950910348006
    H = objMap.Height * 0.566950910348006
    
    sglPerW = 1
    sglPerH = 1
    
    If W > cx Then sglPerW = cx / W
    If H > cy Then sglPerH = cy / H
    
    If W > cx Or H > cy Then
        sglPer = IIf(sglPerW > sglPerH, sglPerH, sglPerW)
        W = W * sglPer
        H = H * sglPer
    End If

    GetPicturePrintHeight = H

    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox "不能打开文件(" & strFile & "),该文件可能正在使用或文件不存在!"
'    Resume
End Function

Private Function MakeConclusion(ByRef objRect As USERRECT, ByRef objFont As USERFONT, ByVal str结论描述 As String, ByVal str科室 As String) As Boolean
    '******************************************************************************************************************
    '功能：产生小结数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTmp      As String
    Dim objSvrRect  As USERRECT
    Dim objSvrFont  As USERFONT
    Dim intRows     As Integer
    Dim aryTmp      As Variant
    Dim lngLoop     As Long
    Dim blnBold     As Boolean
    Dim blnItatic   As Boolean
    
    blnBold = objFont.Bold
    blnItatic = objFont.Italic
    
    aryTmp = Split(str结论描述, "'")
            
    If objRect.Y1 + 2 * objRect.B0 + 1 * (gclsLisReportLib.GetTextHeight("高") + objRect.R0) - objRect.R0 > gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - gclsLisReportLib.Paper.PageFoot - gclsLisReportLib.Paper.SpaceBottom Then
    
        '超过了本页高度
        
        Call ContinuePage(objRect, objFont)
        
        objFont.Bold = True
        objFont.Italic = False
        objFont.BackColor = USERCOLOR.白色
        objFont.ForeColor = USERCOLOR.黑色
        objFont.Underline = False
        gclsLisReportLib.Font = objFont
        
        objRect.Y1 = 0
        objRect.X0 = gclsLisReportLib.Paper.BorderLeft
        objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
'        Call AppendPrintData(rsData, "项目", "文本", objRect, objFont, gclsLisReportLib.Paper, , , str科室)
        objFont.MuliLine = True
        Call mclsLisReportData.InsertText(str科室, objRect, objFont)
        objFont.MuliLine = False
        objFont.Bold = blnBold
        objFont.Italic = blnItatic
    End If
                    
    objFont.Bold = True
    objFont.ForeColor = USERCOLOR.黑色
    
    Call SaveRect(objRect, objSvrRect)
    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.Y0 = objRect.Y1 + 60
    objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(1.5)
    objRect.Y1 = 0
    
    Call mclsLisReportData.InsertText(Mid(str科室, 6) & "小结：", objRect, objFont)
    
    Call SaveRect(objSvrRect, objRect)
    For lngLoop = 0 To UBound(aryTmp)
        strTmp = IIf(UBound(aryTmp) > 0, lngLoop + 1 & ".", "") & aryTmp(lngLoop)

        If objRect.Y0 + 2 * objRect.B0 + (gclsLisReportLib.GetTextHeight("高") + objRect.R0) - objRect.R0 > gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - gclsLisReportLib.Paper.PageFoot - gclsLisReportLib.Paper.SpaceBottom Then
            '换页
            Call ContinuePage(objRect, objFont, False)
            
            objFont.Bold = True
            objRect.X0 = gclsLisReportLib.Paper.BorderLeft
            objRect.Y1 = 0
            Call mclsLisReportData.InsertText(Mid(str科室, 6) & "小结：", objRect, objFont)
            objRect.Y1 = objRect.Y0 - 60
            
        End If
        
        objFont.Bold = False
        objFont.ForeColor = USERCOLOR.黑色
        objRect.X0 = gclsLisReportLib.Paper.BorderLeft + gclsLisReportLib.GetTextWidth(Mid(str科室, 6) & "小结： ")
        objRect.Y0 = objRect.Y1 + 60
        objRect.X1 = objRect.X0 + gclsLisReportLib.GetTwipsX(mlngGridColWidth(1) + mlngGridColWidth(2) + mlngGridColWidth(3) + mlngGridColWidth(4) + mlngGridColWidth(5) - 1.5)
        objRect.Y1 = 0
        objFont.MuliLine = True
        
        Call mclsLisReportData.InsertText(strTmp, objRect, objFont)
        objFont.MuliLine = False
    Next
    
    objFont.Bold = blnBold
    objFont.Italic = blnItatic
    
    MakeConclusion = True
    
End Function

Private Function ContinuePage(ByRef objRect As USERRECT, ByRef objFont As USERFONT, Optional ByVal blnLine As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：续页处理
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objSvrRect  As USERRECT
    Dim objSvrFont  As USERFONT
    Dim blnBold     As Boolean
    Dim blnItatic   As Boolean
    
    blnBold = objFont.Bold
    blnItatic = objFont.Italic
    
    Call SaveRect(objRect, objSvrRect)
    Call SaveFont(objFont, objSvrFont)
    
    objFont.Bold = False
    objFont.BackColor = USERCOLOR.白色
    objFont.ForeColor = USERCOLOR.黑色
    objFont.Underline = False
    objFont.Italic = False
    gclsLisReportLib.Font = objFont
                    
    '补线
    objFont.ForeColor = USERCOLOR.表格线色
    objFont.Size = 9
    
    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.Y0 = objRect.Y1
    objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
    objRect.Y1 = objRect.Y0
    If blnLine Then Call mclsLisReportData.InsertLine(objRect, objFont)
    
    '补续下页阅读标志
    objFont.ForeColor = 0
    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.Y0 = gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.PageFoot - gclsLisReportLib.Paper.BorderBottom - gclsLisReportLib.GetTwipsY(0.4)
    objRect.Y1 = 0
    Call mclsLisReportData.InsertContinueText("(续下页)", objRect, objFont)
    
    mclsLisReportData.NewPage
    objRect.Page = objRect.Page + 1
    
    '补继上页阅读标志
    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.Y0 = gclsLisReportLib.Paper.BorderTop + gclsLisReportLib.Paper.PageHead + gclsLisReportLib.GetTwipsY(0.1)
    objRect.Y1 = 0
    Call mclsLisReportData.InsertContinueText("(继上页)", objRect, objFont)
    
    '补线
    objFont.ForeColor = USERCOLOR.表格线色
    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.Y0 = gclsLisReportLib.Paper.BorderTop + gclsLisReportLib.Paper.PageHead + gclsLisReportLib.Paper.SpaceTop
    objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
    objRect.Y1 = objRect.Y0
    If blnLine Then Call mclsLisReportData.InsertLine(objRect, objFont)
    
    objRect.Y0 = objRect.Y1 + gclsLisReportLib.GetTextHeight("高") + objRect.R0
    objRect.Y1 = 0
    
    objSvrRect.Y0 = objRect.Y0
    objSvrRect.Y1 = objRect.Y1
    objSvrRect.Page = objRect.Page
    objFont.Bold = blnBold
    objFont.Italic = blnItatic
    
    Call SaveRect(objSvrRect, objRect)
    Call SaveFont(objSvrFont, objFont)
        
End Function

Private Function Get_Text(ByVal strText As String, ByVal lngWidth As Long, ByVal lngSpace As Long) As String
    Dim strLineTmp As String
    Dim strLine As String
    Dim strTmp As String
    
    strLineTmp = strText
    strTmp = ""
    strLine = ""
    
    If gclsLisReportLib.GetTextWidth(strLineTmp) > lngWidth Or InStr(strLineTmp, vbCrLf) > 0 Then
        Do While strLineTmp <> ""
            If (Mid(strLineTmp, 1, 1) = Chr(10) Or Mid(strLineTmp, 1, 1) = Chr(13)) Then
                If strLine <> "" Then strTmp = strTmp & strLine & vbCrLf & Space(lngSpace)
                strLine = ""
            ElseIf gclsLisReportLib.GetTextWidth(strLine & Mid(strLineTmp, 1, 1)) > lngWidth Then
                strTmp = strTmp & strLine & vbCrLf & Space(lngSpace)
                strLine = Mid(strLineTmp, 1, 1)
            Else
                strLine = strLine & Mid(strLineTmp, 1, 1)
            End If
            strLineTmp = Mid(strLineTmp, 2)
        Loop
        If strLine <> "" Then strTmp = strTmp & strLine
    Else
        strTmp = strText
    End If
    Get_Text = strTmp
End Function

Private Function MakeDeviceName(ByRef objRect As USERRECT, ByRef objFont As USERFONT, ByVal rsItem As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：处理每个申请的表头项
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strItemTmp  As String
    
    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.Y1 = objRect.Y1 + 5
    
    strItemTmp = "标 本 号:" & (rsItem!标本序号)   '"仪  器:" & rsItem!仪器
    Call mclsLisReportData.InsertText(strItemTmp, objRect, objFont)
    'strItemTmp = GetSampleItemName(rsItem!标本id)
    strItemTmp = "项  目:" & GetSampleItemName(rsItem!标本id) '"标 本 号:" & (rsItem!标本序号)
    objFont.Size = 10.5
    objRect.X0 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderLeft - gclsLisReportLib.Paper.BorderRight - gclsLisReportLib.GetTextWidth(strItemTmp) - 450
    Call mclsLisReportData.InsertText(strItemTmp, objRect, objFont)
    If Nvl(rsItem!急诊标本) <> "" Then
        objFont.ForeColor = USERCOLOR.红色
        objRect.X0 = objRect.X0 + gclsLisReportLib.GetTextWidth(strItemTmp)
        strItemTmp = "(" & Nvl(rsItem!急诊标本) & ")"
        Call mclsLisReportData.InsertText(strItemTmp, objRect, objFont)
    End If
End Function

Private Function GetSampleItemName(ByVal lng标本id As Long) As String
     Dim rsTemp As New ADODB.Recordset
     Dim strTemp As String
     Dim strSql As String
     strSql = "select 序号,医嘱内容 from ( "
     strSql = strSql & " select  min(b.序号) as 序号 , b.医嘱内容  from 检验项目分布 a,病人医嘱记录 b   where a.医嘱id=b.相关id and   标本id=[1]"
     strSql = strSql & " group by b.医嘱内容"
     strSql = strSql & "  order by 序号 ) where rownum <2 "
     Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, gstrSysName, lng标本id)
     If rsTemp.RecordCount > 0 Then
        If ChkRsState(rsTemp) Then
           strTemp = ""
        Else
           strTemp = Nvl(rsTemp!医嘱内容, "")
        End If
     Else
        strSql = "select 申请项目 from 检验报告记录 where id=[1]"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, gstrSysName, lng标本id)
        If rsTemp.RecordCount > 0 Then
            strTemp = Nvl(rsTemp!申请项目, "")
        End If
     End If
     GetSampleItemName = strTemp
End Function


Private Function MakePatientTitle(ByRef objRect As USERRECT, ByRef objFont As USERFONT, ByVal btyType As Integer, ByVal rsItem As ADODB.Recordset, ByVal strSign As String) As Boolean
    '******************************************************************************************************************
    '功能：处理每个申请的病人信息
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim intLoop         As Integer
    Dim int行间距       As Integer
    Dim varTmp          As Variant
    Dim strTmp          As String
    Dim strPerson       As String
    Dim lngMapHeight    As Single
    Dim objMap          As StdPicture
    
    If btyType = 1 Then
        strPerson = "姓  名:" & Nvl(rsItem!姓名) & "'申请科室:" & Nvl(rsItem!申请科室) & "'申请医生:" & Nvl(rsItem!申请医生) & _
            "'申请时间:" & Nvl(rsItem!申请时间) & "'性  别:" & Nvl(rsItem!性别) & "'床    号:" & Nvl(rsItem!床号) & _
            "'标本类型:" & Nvl(rsItem!标本类型) & "'采样时间:" & Nvl(rsItem!采样时间) & "'年  龄:" & Nvl(rsItem!年龄) & "'" & Nvl(rsItem!住院号) & _
            "'诊    断:" & Nvl(rsItem!临床特征)
    Else
        strPerson = "签收时间:" & Nvl(rsItem!签收时间) & "'报告时间:" & Nvl(rsItem!报告时间) & "'检验医师:" & Nvl(rsItem!报告人) & "'审核医师:" & Nvl(rsItem!审核人)
    End If

    int行间距 = mclsLisReportLib.GetTwipsY(0.14)
    varTmp = Split(strPerson, "'")
    
    For intLoop = 0 To UBound(varTmp)
        strTmp = CStr(varTmp(intLoop))
        If intLoop = 0 And (intLoop Mod 4) = 0 Then
            objRect.X0 = gclsLisReportLib.Paper.BorderLeft
            objRect.X1 = 0
            objRect.Y0 = objRect.Y1 + 30
            objRect.Y1 = 0
        ElseIf intLoop > 0 And (intLoop Mod 4) = 0 Then
            If btyType = 1 Then
                objFont.Bold = False
            Else
                objFont.Bold = True
            End If
            objRect.X0 = gclsLisReportLib.Paper.BorderLeft
            objRect.X1 = 0
            objRect.Y0 = objRect.Y1 + int行间距
            objRect.Y1 = 0
        Else
            objRect.X0 = objRect.X0 + (gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderLeft - gclsLisReportLib.Paper.BorderRight) / 4 - 100
            objRect.X1 = 0
            objRect.Y0 = objRect.Y0
            objRect.Y1 = 0
        End If
        If Len(strSign) > 0 And intLoop = UBound(varTmp) Then
            Call mclsLisReportData.InsertText(Split(strTmp, ":")(0) & ":", objRect, objFont)
            Set objMap = LoadPicture(strSign)
            objRect.X1 = objRect.X0 + mclsLisReportLib.GetTextWidth(Split(strTmp, ":")(0) & ":") + objMap.Width + 100
            objRect.Y0 = objRect.Y0 - 150
            objRect.Y1 = objRect.Y0 + 550
            Call mclsLisReportData.InsertImage(objRect, objFont, strSign)
        Else
            Call mclsLisReportData.InsertText(strTmp, objRect, objFont)
        End If
    Next
End Function

Private Function MakeLisPicResultData(ByRef objRect As USERRECT, ByRef objFont As USERFONT, ByVal strFilename As String) As Boolean
    '******************************************************************************************************************
    '功能：产生图形结果数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTmp          As String
    Dim strFile         As String
    Dim lngMapHeight    As Single
    Dim dblY1           As Double
    Dim bytAngle        As Byte
    Dim intLoop         As Integer
    
    
On Error GoTo ErrH
'    strFilename = "974_DIFF.jpg#974_BASO.jpg#974_RBC.jpg#974_PLT.jpg" '
    For intLoop = 0 To UBound(Split(strFilename, "#"))
        strFile = gstrFilePath & "\" & Split(strFilename, "#")(intLoop)
        If strFile <> "" Then
    
            lngMapHeight = GetPicPrintHeight(strFile, gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight, 6000, 1)
'            If lngMapHeight = 0 Then lngMapHeight = 5292
            If lngMapHeight > gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - gclsLisReportLib.Paper.PageFoot - gclsLisReportLib.Paper.SpaceBottom Then lngMapHeight = gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - gclsLisReportLib.Paper.BorderTop - gclsLisReportLib.GetTwipsY(0.5)
            
            If intLoop = 0 And objRect.Y1 + 2 * objRect.B0 + lngMapHeight - 2292 - objRect.R0 > gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - gclsLisReportLib.Paper.PageFoot - gclsLisReportLib.Paper.SpaceBottom Then
    
               '换页
                mclsLisReportData.NewPage
                objRect.Page = objRect.Page + 1
    
'               objRect.X0 = gclsLisReportLib.Paper.BorderLeft
'               objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
'               objRect.Y0 = gclsLisReportLib.Paper.BorderTop + gclsLisReportLib.Paper.PageHead + gclsLisReportLib.GetTwipsY(0.5)
'               objRect.Y1 = lngMapHeight
'               If intLoop = 0 Then
'                    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
                objRect.X1 = lngMapHeight - 2292
'                    objRect.Y0 = objRect.Y1 + 20
'                    objRect.Y1 = objRect.Y0 + lngMapHeight - 2292
                objRect.Y0 = gclsLisReportLib.Paper.PageHead + gclsLisReportLib.Paper.BorderTop + gclsLisReportLib.GetTwipsY(0.2) '- gclsLisReportLib.GetTwipsY(0.2) '
                objRect.Y1 = objRect.Y0 + lngMapHeight - 2292
'                Else
'                    objRect.X0 = objRect.X0 + lngMapHeight - 2292
'                    objRect.X1 = objRect.X0 + lngMapHeight - 2292
'                    If intLoop = 1 Then
'                        objRect.Y1 = objRect.Y1 - 140
'                    Else
'                        objRect.Y1 = objRect.Y1 + 140
'                    End If
'                End If
            Else
                If intLoop = 0 Then
                    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
                    objRect.X1 = lngMapHeight - 2292
                    objRect.Y0 = objRect.Y1 + 20
                    objRect.Y1 = objRect.Y0 + lngMapHeight - 2292
                Else
                    objRect.X0 = objRect.X0 + lngMapHeight - 2292
                    objRect.X1 = objRect.X0 + lngMapHeight - 2292
                    If intLoop = 1 Then
                        objRect.Y1 = objRect.Y1 - 140
                    Else
                        objRect.Y1 = objRect.Y1 + 140
                    End If
'                    objRect.Y0 = objRect.Y0 + 20
'                    objRect.Y1 = objRect.Y0 + lngMapHeight - 2292
                End If
            End If
            bytAngle = 0

            Call mclsLisReportData.InsertImage(objRect, objFont, strFile, 1, 1)
        End If
    Next
    
    MakeLisPicResultData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetPicPrintHeight(ByVal strFile As String, ByVal cx As Double, ByVal cy As Double, ByVal lngAngle As Long) As Single
    '******************************************************************************************************************
    '功能：根据相册大小自动等比例缩放相片文件
    '参数：缩放前的相片文件
    '返回：缩放后相片的高度
    '******************************************************************************************************************
    Dim strTmp      As String
    Dim objMap      As StdPicture
    Dim W           As Single
    Dim H           As Single
    Dim sglPerW     As Single
    Dim sglPerH     As Single
    Dim sglPer      As Single
    Dim X           As Long
    Dim Y           As Long
    
On Error GoTo ErrH
    
    If strFile = "" Then Exit Function
    
    Set objMap = VB.LoadPicture(strFile)
    
    W = objMap.Width
    H = objMap.Height

    If lngAngle = 1 Then
        '横向输出
        If H > cx Then
            H = W * (cx / H)
        End If
    Else
        '纵向输出(默认)
        If W > cx Then
            H = H * (cx / W)
        End If
    End If
    
'
    GetPicPrintHeight = H

    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrH:
    err.Clear
    Call ErrLog("计算图像高度", "", "不能打开文件(" & strFile & "),该文件可能正在使用或文件不存在!", "")
'    Resume
End Function

'类中自带的事件
'######################################################################################################################
Private Sub Class_Initialize()
    
End Sub

Private Sub Class_Terminate()
    Set mrsParameter = Nothing
    Set gcnOracle = Nothing
    Set mclsLisPackage = Nothing
    Set mclsLisReportData = Nothing
    Set mclsLisReportLib = Nothing
    Set mclsDrawLisImage = Nothing
End Sub
