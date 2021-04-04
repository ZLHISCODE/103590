Attribute VB_Name = "mdlCISOutPath"
Option Explicit

Public Function ExportOutPathToXML(ByVal lng路径ID As Long, ByVal int版本号 As Integer, ByVal strFile As String) As Boolean
'功能：导出门诊临床路径到XML文件
'参数：strFile=包含路径的文件名
'说明：导出包含路径信息和指定版本的信息
    Dim xPath As DOMDocument
    Dim xRoot As IXMLDOMElement
    Dim xNode As IXMLDOMNode
    Dim xSubNode1 As IXMLDOMNode
    Dim xSubNode2 As IXMLDOMNode
    Dim xSubNode3 As IXMLDOMNode
    Dim xSubNode4 As IXMLDOMNode
    Dim xSubNode5 As IXMLDOMNode
    Dim xPI As IXMLDOMProcessingInstruction
    
    Dim rsTmp As ADODB.Recordset
    Dim rsClone As ADODB.Recordset
    Dim rsItem As ADODB.Recordset
    Dim rsItemAdvice As ADODB.Recordset
    Dim rsItemEPR As ADODB.Recordset
    Dim rsEvalMark As ADODB.Recordset
    Dim rsEvalCond As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    Set xPath = New DOMDocument
    
    '注释
    xPath.appendChild xPath.createComment(gstrSysName & "  操作员:" & UserInfo.姓名 & ",部门:" & UserInfo.部门名 & ",时间:" & Format(Now(), "yyyy-MM-dd HH:mm:ss"))
    
    '根结点
    Set xRoot = xPath.createElement("ClinicalPathways")
    Set xPath.documentElement = xRoot
    Call xRoot.setAttribute("ID", lng路径ID)
    Call xRoot.setAttribute("Version", int版本号)

    '门诊临床路径信息
    strSql = "Select A.分类,A.编码,A.名称,A.通用,A.最新版本," & _
        " A.适用性别,A.适用年龄,A.说明,B.标准治疗时间,B.标准费用," & _
        " B.版本说明,B.创建人,B.创建时间,B.审核人,B.审核时间,B.停用人,B.停用时间,A.最大间隔时间 " & _
        " From 门诊路径目录 A,门诊路径版本 B Where A.ID=B.路径ID And A.ID=[1] And B.版本号=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng路径ID, int版本号)
    
    Set xNode = CreateNode(1, xRoot, "PathInfo", NODE_ELEMENT, "")
        CreateNode 2, xNode, "分类", , rsTmp!分类
        CreateNode 2, xNode, "编码", , rsTmp!编码
        CreateNode 2, xNode, "名称", , rsTmp!名称
        CreateNode 2, xNode, "通用", , NVL(rsTmp!通用)
        CreateNode 2, xNode, "最新版本", , NVL(rsTmp!最新版本)
        CreateNode 2, xNode, "适用性别", , NVL(rsTmp!适用性别)
        CreateNode 2, xNode, "适用年龄", , NVL(rsTmp!适用年龄)
        CreateNode 2, xNode, "说明", , NVL(rsTmp!说明)
        CreateNode 2, xNode, "标准治疗时间", , NVL(rsTmp!标准治疗时间)
        CreateNode 2, xNode, "标准费用", , NVL(rsTmp!标准费用)
        CreateNode 2, xNode, "版本说明", , NVL(rsTmp!版本说明)
        CreateNode 2, xNode, "创建人", , NVL(rsTmp!创建人)
        CreateNode 2, xNode, "创建时间", , Format(NVL(rsTmp!创建时间), "yyyy-MM-dd HH:mm:ss")
        CreateNode 2, xNode, "审核人", , NVL(rsTmp!审核人)
        CreateNode 2, xNode, "审核时间", , Format(NVL(rsTmp!审核时间), "yyyy-MM-dd HH:mm:ss")
        CreateNode 2, xNode, "停用人", , NVL(rsTmp!停用人)
        CreateNode 2, xNode, "停用时间", , Format(NVL(rsTmp!停用时间), "yyyy-MM-dd HH:mm:ss")
        CreateNode 2, xNode, "最大间隔时间", , NVL(rsTmp!最大间隔时间)
    '门诊路径科室
    strSql = "Select B.ID,B.编码,B.名称 From 门诊路径科室 A,部门表 B Where A.路径ID=[1] And A.科室ID=B.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng路径ID)
    If Not rsTmp.EOF Then
        Set xNode = CreateNode(1, xRoot, "PathDepts", NODE_ELEMENT, "")
        Do While Not rsTmp.EOF
            Set xSubNode1 = CreateNode(2, xNode, "PathDept", NODE_ELEMENT, "")
                CreateNode 3, xSubNode1, "科室ID", , rsTmp!ID
                CreateNode 3, xSubNode1, "编码", , rsTmp!编码
                CreateNode 3, xSubNode1, "名称", , rsTmp!名称
            rsTmp.MoveNext
        Loop
    End If
    
    '门诊路径病种
    strSql = "Select A.疾病ID,B.编码 as 疾病码,B.名称 as 疾病名," & _
        " A.诊断ID,C.编码 as 诊断码,C.名称 as 诊断名 " & _
        " From 门诊路径病种 A,疾病编码目录 B,疾病诊断目录 C" & _
        " Where Nvl(A.疾病ID,0)=B.ID(+) And Nvl(A.诊断ID,0)=C.ID(+) And A.路径ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng路径ID)
    If Not rsTmp.EOF Then
        Set xNode = CreateNode(1, xRoot, "PathDiseases", NODE_ELEMENT, "")
        Do While Not rsTmp.EOF
            Set xSubNode1 = CreateNode(2, xNode, "PathDisease", NODE_ELEMENT, "")
                CreateNode 3, xSubNode1, "疾病ID", , NVL(rsTmp!疾病id)
                CreateNode 3, xSubNode1, "疾病码", , NVL(rsTmp!疾病码)
                CreateNode 3, xSubNode1, "疾病名", , NVL(rsTmp!疾病名)
                CreateNode 3, xSubNode1, "诊断ID", , NVL(rsTmp!诊断id)
                CreateNode 3, xSubNode1, "诊断码", , NVL(rsTmp!诊断码)
                CreateNode 3, xSubNode1, "诊断名", , NVL(rsTmp!诊断名)
            rsTmp.MoveNext
        Loop
    End If
    
    '导入评估
    strSql = "Select B.评估类型,B.阶段ID,A.ID,A.评估指标,A.指标类型,A.指标结果" & _
        " From 门诊路径评估指标 A,门诊路径评估 B" & _
        " Where A.评估ID=B.ID And B.路径ID=[1] And 版本号=[2]" & _
        " Order by B.评估类型,B.阶段ID,A.序号"
    Set rsEvalMark = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng路径ID, int版本号)
    
    strSql = "Select B.评估类型,B.阶段ID,A.指标ID,A.项目ID,A.关系式,A.条件值,A.条件组合" & _
        " From 门诊路径评估条件 A,门诊路径评估 B" & _
        " Where A.评估ID=B.ID And B.路径ID=[1] And 版本号=[2]" & _
        " Order by B.评估类型,B.阶段ID"
    Set rsEvalCond = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng路径ID, int版本号)
    
    rsEvalMark.Filter = "评估类型=1"
    rsEvalCond.Filter = "评估类型=1"
    If Not rsEvalMark.EOF Or Not rsEvalCond.EOF Then
        Set xNode = CreateNode(1, xRoot, "ImportEval", NODE_ELEMENT, "")
            If Not rsEvalMark.EOF Then
                Set xSubNode1 = CreateNode(2, xNode, "Marks", NODE_ELEMENT, "")
                Do While Not rsEvalMark.EOF
                    Set xSubNode2 = CreateNode(3, xSubNode1, "Mark", NODE_ELEMENT, "")
                        CreateNode 4, xSubNode2, "ID", , rsEvalMark!ID
                        CreateNode 4, xSubNode2, "评估指标", , rsEvalMark!评估指标
                        CreateNode 4, xSubNode2, "指标类型", , rsEvalMark!指标类型
                        CreateNode 4, xSubNode2, "指标结果", , rsEvalMark!指标结果
                    rsEvalMark.MoveNext
                Loop
            End If
            If Not rsEvalCond.EOF Then
                Set xSubNode1 = CreateNode(2, xNode, "Conditions", NODE_ELEMENT, "")
                Do While Not rsEvalCond.EOF
                    Set xSubNode2 = CreateNode(3, xSubNode1, "Condition", NODE_ELEMENT, "")
                        CreateNode 4, xSubNode2, "指标ID", , rsEvalCond!指标ID
                        CreateNode 4, xSubNode2, "关系式", , rsEvalCond!关系式
                        CreateNode 4, xSubNode2, "条件值", , rsEvalCond!条件值
                        CreateNode 4, xSubNode2, "条件组合", , rsEvalCond!条件组合
                    rsEvalCond.MoveNext
                Loop
            End If
    End If
    
    '门诊路径医嘱内容
    strSql = "Select Distinct A.ID,A.相关ID,A.序号,A.期效,A.诊疗项目ID,D.编码 as 诊疗编码,D.名称 as 诊疗名称," & _
        " A.收费细目ID,E.编码 as 收费编码,E.名称 as 收费名称,A.医嘱内容,A.单次用量,A.总给予量," & _
        " A.标本部位,A.检查方法,A.医生嘱托,A.执行频次,A.频率次数,A.频率间隔,A.间隔单位," & _
        " A.执行性质,A.执行科室ID,F.编码 as 执行科室码,F.名称 as 执行科室名,A.时间方案,A.是否缺省,A.是否备选,A.配方ID,A.组合项目ID" & _
        " From 门诊路径医嘱内容 A,门诊路径医嘱 B,门诊路径项目 C,诊疗项目目录 D,收费项目目录 E,部门表 F" & _
        " Where A.ID=B.医嘱内容ID And B.路径项目ID=C.ID And C.路径ID=[1] And C.版本号=[2]" & _
        " And Nvl(A.诊疗项目ID,0)=D.ID(+) And Nvl(A.收费细目ID,0)=E.ID(+) And Nvl(A.执行科室ID,0)=F.ID(+)" & _
        " Order by A.序号,A.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng路径ID, int版本号)
    If Not rsTmp.EOF Then
        Set xNode = CreateNode(1, xRoot, "PathAdvices", NODE_ELEMENT, "")
        Do While Not rsTmp.EOF
            Set xSubNode1 = CreateNode(2, xNode, "PathAdvice", NODE_ELEMENT, "")
                CreateNode 3, xSubNode1, "ID", , rsTmp!ID
                CreateNode 3, xSubNode1, "相关ID", , NVL(rsTmp!相关id)
                CreateNode 3, xSubNode1, "序号", , rsTmp!序号
                CreateNode 3, xSubNode1, "期效", , rsTmp!期效
                CreateNode 3, xSubNode1, "诊疗项目ID", , NVL(rsTmp!诊疗项目ID)
                CreateNode 3, xSubNode1, "诊疗编码", , NVL(rsTmp!诊疗编码)
                CreateNode 3, xSubNode1, "诊疗名称", , NVL(rsTmp!诊疗名称)
                CreateNode 3, xSubNode1, "收费细目ID", , NVL(rsTmp!收费细目ID)
                CreateNode 3, xSubNode1, "收费编码", , NVL(rsTmp!收费编码)
                CreateNode 3, xSubNode1, "收费名称", , NVL(rsTmp!收费名称)
                CreateNode 3, xSubNode1, "医嘱内容", , NVL(rsTmp!医嘱内容)
                CreateNode 3, xSubNode1, "单次用量", , NVL(rsTmp!单次用量)
                CreateNode 3, xSubNode1, "总给予量", , NVL(rsTmp!总给予量)
                CreateNode 3, xSubNode1, "标本部位", , NVL(rsTmp!标本部位)
                CreateNode 3, xSubNode1, "检查方法", , NVL(rsTmp!检查方法)
                CreateNode 3, xSubNode1, "医生嘱托", , NVL(rsTmp!医生嘱托)
                CreateNode 3, xSubNode1, "执行频次", , NVL(rsTmp!执行频次)
                CreateNode 3, xSubNode1, "频率次数", , NVL(rsTmp!频率次数)
                CreateNode 3, xSubNode1, "频率间隔", , NVL(rsTmp!频率间隔)
                CreateNode 3, xSubNode1, "间隔单位", , NVL(rsTmp!间隔单位)
                CreateNode 3, xSubNode1, "执行性质", , NVL(rsTmp!执行性质)
                CreateNode 3, xSubNode1, "执行科室ID", , NVL(rsTmp!执行科室ID)
                CreateNode 3, xSubNode1, "执行科室码", , NVL(rsTmp!执行科室码)
                CreateNode 3, xSubNode1, "执行科室名", , NVL(rsTmp!执行科室名)
                CreateNode 3, xSubNode1, "时间方案", , NVL(rsTmp!时间方案)
                CreateNode 3, xSubNode1, "是否缺省", , NVL(rsTmp!是否缺省, 0)
                CreateNode 3, xSubNode1, "是否备选", , NVL(rsTmp!是否备选, 0)
                CreateNode 3, xSubNode1, "配方ID", , NVL(rsTmp!配方ID)
                CreateNode 3, xSubNode1, "组合项目ID", , NVL(rsTmp!组合项目ID)
            rsTmp.MoveNext
        Loop
    End If
    
    '门诊路径分类
    strSql = "Select 名称 From 门诊路径分类 Where 路径ID=[1] And 版本号=[2] Order by 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng路径ID, int版本号)
    
    Set xNode = CreateNode(1, xRoot, "PathCategorys", NODE_ELEMENT, "")
    Do While Not rsTmp.EOF
        Set xSubNode1 = CreateNode(2, xNode, "PathCategory", NODE_ELEMENT, NVL(rsTmp!名称))
        CreateNode 2, xSubNode1, "名称", NODE_ELEMENT, NVL(rsTmp!名称)
        rsTmp.MoveNext
    Loop
    
    '门诊路径阶段/项目
    strSql = "Select ID,Nvl(父ID,0) as 父ID,序号,名称,开始天数,结束天数,分类,说明" & _
        " From 门诊路径阶段 Where 路径ID=[1] And 版本号=[2] Order by Nvl(父ID,0) Desc,序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng路径ID, int版本号)
    
    strSql = "Select ID,阶段ID,分类,项目序号,项目内容,执行方式,项目结果,图标ID,内容要求" & _
        " From 门诊路径项目 Where 路径ID=[1] And 版本号=[2] Order by 阶段ID,分类,项目序号"
    Set rsItem = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng路径ID, int版本号)
    
    strSql = "Select A.路径项目ID,A.医嘱内容ID From 门诊路径医嘱 A,门诊路径项目 B" & _
        " Where A.路径项目ID=B.ID And B.路径ID=[1] And 版本号=[2]"
    Set rsItemAdvice = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng路径ID, int版本号)
    
    strSql = "Select A.项目ID,A.文件ID,C.编号,C.名称 From 门诊路径病历 A,门诊路径项目 B,病历文件列表 C" & _
        " Where A.项目ID=B.ID And A.文件ID=C.ID And B.路径ID=[1] And 版本号=[2]"
    Set rsItemEPR = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng路径ID, int版本号)
    
    Set rsClone = rsTmp.Clone: rsTmp.Filter = "父ID=0"
    
    Set xNode = CreateNode(1, xRoot, "PathTimeSteps", NODE_ELEMENT, "")
    Do While Not rsTmp.EOF
        '缺省分支
        Set xSubNode1 = CreateNode(2, xNode, "PathTimeStep", NODE_ELEMENT, "")
            CreateNode 3, xSubNode1, "ID", , rsTmp!ID
            CreateNode 3, xSubNode1, "父ID", , ""
            CreateNode 3, xSubNode1, "序号", , rsTmp!序号
            CreateNode 3, xSubNode1, "名称", , rsTmp!名称
            CreateNode 3, xSubNode1, "开始天数", , NVL(rsTmp!开始天数)
            CreateNode 3, xSubNode1, "结束天数", , NVL(rsTmp!结束天数)
            CreateNode 3, xSubNode1, "说明", , NVL(rsTmp!说明)
            CreateNode 3, xSubNode1, "分类", , NVL(rsTmp!分类)
            
            '阶段的项目
            rsItem.Filter = "阶段ID=" & rsTmp!ID
            Set xSubNode2 = CreateNode(3, xSubNode1, "Items", NODE_ELEMENT, "")
            Do While Not rsItem.EOF
                Set xSubNode3 = CreateNode(4, xSubNode2, "Item", NODE_ELEMENT, "")
                    CreateNode 5, xSubNode3, "ID", , rsItem!ID
                    CreateNode 5, xSubNode3, "分类", , rsItem!分类
                    CreateNode 5, xSubNode3, "项目序号", , rsItem!项目序号
                    CreateNode 5, xSubNode3, "项目内容", , rsItem!项目内容
                    CreateNode 5, xSubNode3, "执行方式", , NVL(rsItem!执行方式)
                    CreateNode 5, xSubNode3, "项目结果", , NVL(rsItem!项目结果)
                    CreateNode 5, xSubNode3, "图标ID", , NVL(rsItem!图标ID)
                    CreateNode 5, xSubNode3, "内容要求", , NVL(rsItem!内容要求, 0)

                    '项目对应的医嘱
                    rsItemAdvice.Filter = "路径项目ID=" & rsItem!ID
                    If Not rsItemAdvice.EOF Then
                        Set xSubNode4 = CreateNode(5, xSubNode3, "Advices", NODE_ELEMENT, "")
                        Do While Not rsItemAdvice.EOF
                            CreateNode 6, xSubNode4, "Advice", , rsItemAdvice!医嘱内容ID
                            rsItemAdvice.MoveNext
                        Loop
                    End If
                    '项目对应的病历
                    rsItemEPR.Filter = "项目ID=" & rsItem!ID
                    If Not rsItemEPR.EOF Then
                        Set xSubNode4 = CreateNode(5, xSubNode3, "EPRFiles", NODE_ELEMENT, "")
                        Do While Not rsItemEPR.EOF
                            Set xSubNode5 = CreateNode(6, xSubNode4, "EPRFile", NODE_ELEMENT, "")
                                CreateNode 7, xSubNode5, "文件ID", , rsItemEPR!文件ID
                                CreateNode 7, xSubNode5, "文件编号", , rsItemEPR!编号
                                CreateNode 7, xSubNode5, "文件名称", , rsItemEPR!名称
                            rsItemEPR.MoveNext
                        Loop
                    End If
                    
                rsItem.MoveNext
            Loop
        
            '阶段的评估
            rsEvalMark.Filter = "评估类型=2 And 阶段ID=" & rsTmp!ID
            rsEvalCond.Filter = "评估类型=2 And 阶段ID=" & rsTmp!ID
            If Not rsEvalMark.EOF Or Not rsEvalCond.EOF Then
                Set xSubNode2 = CreateNode(3, xSubNode1, "StepEval", NODE_ELEMENT, "")
                    If Not rsEvalMark.EOF Then
                        Set xSubNode3 = CreateNode(4, xSubNode2, "Marks", NODE_ELEMENT, "")
                        Do While Not rsEvalMark.EOF
                            Set xSubNode4 = CreateNode(5, xSubNode3, "Mark", NODE_ELEMENT, "")
                                CreateNode 6, xSubNode4, "ID", , rsEvalMark!ID
                                CreateNode 6, xSubNode4, "评估指标", , rsEvalMark!评估指标
                                CreateNode 6, xSubNode4, "指标类型", , rsEvalMark!指标类型
                                CreateNode 6, xSubNode4, "指标结果", , rsEvalMark!指标结果
                            rsEvalMark.MoveNext
                        Loop
                    End If
                    If Not rsEvalCond.EOF Then
                        Set xSubNode3 = CreateNode(4, xSubNode2, "Conditions", NODE_ELEMENT, "")
                        Do While Not rsEvalCond.EOF
                            Set xSubNode4 = CreateNode(5, xSubNode3, "Condition", NODE_ELEMENT, "")
                                CreateNode 6, xSubNode4, "指标ID", , NVL(rsEvalCond!指标ID)
                                CreateNode 6, xSubNode4, "项目ID", , NVL(rsEvalCond!项目ID)
                                CreateNode 6, xSubNode4, "关系式", , rsEvalCond!关系式
                                CreateNode 6, xSubNode4, "条件值", , rsEvalCond!条件值
                                CreateNode 6, xSubNode4, "条件组合", , rsEvalCond!条件组合
                            rsEvalCond.MoveNext
                        Loop
                    End If
            End If
        
        '备选分支
        rsClone.Filter = "父ID=" & rsTmp!ID
        If Not rsClone.EOF Then
            Do While Not rsClone.EOF
                Set xSubNode1 = CreateNode(2, xNode, "PathTimeStep", NODE_ELEMENT, "")
                    CreateNode 3, xSubNode1, "ID", , rsClone!ID
                    CreateNode 3, xSubNode1, "父ID", , rsClone!父ID
                    CreateNode 3, xSubNode1, "序号", , rsClone!序号
                    CreateNode 3, xSubNode1, "名称", , rsClone!名称
                    CreateNode 3, xSubNode1, "开始天数", , NVL(rsClone!开始天数)
                    CreateNode 3, xSubNode1, "结束天数", , NVL(rsClone!结束天数)
                    CreateNode 3, xSubNode1, "说明", , NVL(rsClone!说明)
                
                    '阶段的项目
                    rsItem.Filter = "阶段ID=" & rsClone!ID
                    Set xSubNode2 = CreateNode(3, xSubNode1, "Items", NODE_ELEMENT, "")
                    Do While Not rsItem.EOF
                        Set xSubNode3 = CreateNode(4, xSubNode2, "Item", NODE_ELEMENT, "")
                            CreateNode 5, xSubNode3, "ID", , rsItem!ID
                            CreateNode 5, xSubNode3, "分类", , rsItem!分类
                            CreateNode 5, xSubNode3, "项目序号", , rsItem!项目序号
                            CreateNode 5, xSubNode3, "项目内容", , rsItem!项目内容
                            CreateNode 5, xSubNode3, "执行方式", , NVL(rsItem!执行方式)
                            CreateNode 5, xSubNode3, "项目结果", , NVL(rsItem!项目结果)
                            CreateNode 5, xSubNode3, "图标ID", , NVL(rsItem!图标ID)
                            
                            '项目对应的医嘱
                            rsItemAdvice.Filter = "路径项目ID=" & rsItem!ID
                            If Not rsItemAdvice.EOF Then
                                Set xSubNode4 = CreateNode(5, xSubNode3, "Advices", NODE_ELEMENT, "")
                                Do While Not rsItemAdvice.EOF
                                    CreateNode 6, xSubNode4, "Advice", , rsItemAdvice!医嘱内容ID
                                    rsItemAdvice.MoveNext
                                Loop
                            End If
                            '项目对应的病历
                            rsItemEPR.Filter = "项目ID=" & rsItem!ID
                            If Not rsItemEPR.EOF Then
                                Set xSubNode4 = CreateNode(5, xSubNode3, "EPRFiles", NODE_ELEMENT, "")
                                Do While Not rsItemEPR.EOF
                                    Set xSubNode5 = CreateNode(6, xSubNode4, "EPRFile", NODE_ELEMENT, "")
                                        CreateNode 7, xSubNode5, "文件ID", , rsItemEPR!文件ID
                                        CreateNode 7, xSubNode5, "文件编号", , rsItemEPR!编号
                                        CreateNode 7, xSubNode5, "文件名称", , rsItemEPR!名称
                                    rsItemEPR.MoveNext
                                Loop
                            End If
                            
                        rsItem.MoveNext
                    Loop
                    
                    '阶段的评估
                    rsEvalMark.Filter = "评估类型=2 And 阶段ID=" & rsClone!ID
                    rsEvalCond.Filter = "评估类型=2 And 阶段ID=" & rsClone!ID
                    If Not rsEvalMark.EOF Or Not rsEvalCond.EOF Then
                        Set xSubNode2 = CreateNode(3, xSubNode1, "StepEval", NODE_ELEMENT, "")
                            If Not rsEvalMark.EOF Then
                                Set xSubNode3 = CreateNode(4, xSubNode2, "Marks", NODE_ELEMENT, "")
                                Do While Not rsEvalMark.EOF
                                    Set xSubNode4 = CreateNode(5, xSubNode3, "Mark", NODE_ELEMENT, "")
                                        CreateNode 6, xSubNode4, "ID", , rsEvalMark!ID
                                        CreateNode 6, xSubNode4, "评估指标", , rsEvalMark!评估指标
                                        CreateNode 6, xSubNode4, "指标类型", , rsEvalMark!指标类型
                                        CreateNode 6, xSubNode4, "指标结果", , rsEvalMark!指标结果
                                    rsEvalMark.MoveNext
                                Loop
                            End If
                            If Not rsEvalCond.EOF Then
                                Set xSubNode3 = CreateNode(4, xSubNode2, "Conditions", NODE_ELEMENT, "")
                                Do While Not rsEvalCond.EOF
                                    Set xSubNode4 = CreateNode(5, xSubNode3, "Condition", NODE_ELEMENT, "")
                                        CreateNode 6, xSubNode4, "指标ID", , NVL(rsEvalCond!指标ID)
                                        CreateNode 6, xSubNode4, "项目ID", , NVL(rsEvalCond!项目ID)
                                        CreateNode 6, xSubNode4, "关系式", , rsEvalCond!关系式
                                        CreateNode 6, xSubNode4, "条件值", , rsEvalCond!条件值
                                        CreateNode 6, xSubNode4, "条件组合", , rsEvalCond!条件组合
                                    rsEvalCond.MoveNext
                                Loop
                            End If
                    End If
                
                rsClone.MoveNext
            Loop
        End If
        
        rsTmp.MoveNext
    Loop
    
    'XML信息
    Set xPI = xPath.createProcessingInstruction("xml", "version='1.0' encoding='gb2312'")
    Call xPath.insertBefore(xPI, xPath.childNodes(0))
    
    '保存成文件
    xPath.Save strFile
    Set xPath = Nothing
    
    ExportOutPathToXML = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set xPath = Nothing
End Function

Public Function ImportOutPathFromXML(ByVal strFile As String, _
    Optional ByVal lng路径ID As Long, Optional ByVal int版本号 As Integer, _
    Optional ByVal intLimit As Integer, Optional ByRef blnLimit As Boolean) As Boolean
'功能：导入指定的门诊临床路径XML文件
'参数：lng路径ID,int版本号=如果指定，则只导入版本相关部分信息；如果没有指定，则根据根据XML中的信息进行路径新增或者完全覆盖
'      intLimit=总体限制的最大路径数量,为0表示不限制
'      blnLimit=是否被允许的最大路径数量所限制导入失败
    Dim rsTmp As ADODB.Recordset
    Dim rsIcon As ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    
    Dim arrSQL As Variant, strSql As String
    Dim colItemID As Collection
    Dim colStepID As Collection
    Dim colMarkID As Collection
    Dim colAdviceID As Collection
    Dim colAdviceOriginalID As Collection
    Dim colBranchID As Collection
    Dim colPreID As Collection
    
    Dim xPath As DOMDocument
    Dim xRoot As IXMLDOMElement
    Dim xNode As IXMLDOMNode
    Dim xSubNode1 As IXMLDOMNode
    Dim xSubNode2 As IXMLDOMNode
    Dim xSubNode3 As IXMLDOMNode
    Dim xSubNode4 As IXMLDOMNode
    Dim xSubNode5 As IXMLDOMNode
    
    Dim str编码 As String, lng阶段ID As Long
    Dim strValue As String, strTemp1 As String
    Dim strTemp2 As String, strTemp3 As String
    Dim blnDo As Boolean, blnTran As Boolean
    Dim i As Long, k As Long, n As Long, m As Long
    Dim strPreStep As String
    Dim strtemp4 As String
    Dim strImportRef As String
    Dim lng导入结果 As Long '记录同一路径项目医嘱的导入状态0，全部未导入，1，全部导入，2，部分导入
    Dim lngCount As Long, str组IDs As String, arrID As Variant, lng组ID As Long, strFilter As String
    Dim lng项目ID As Long
    
    On Error GoTo errH
    
    rsAdvice.Fields.Append "ID", adBigInt
    rsAdvice.Fields.Append "相关ID", adBigInt, , adFldIsNullable
    rsAdvice.Fields.Append "导入参考", adVarChar, 200, adFldIsNullable
    rsAdvice.Fields.Append "项目ID", adBigInt, , adFldIsNullable
    rsAdvice.Fields.Append "导入状态", adInteger
    
    rsAdvice.CursorLocation = adUseClient
    rsAdvice.LockType = adLockOptimistic
    rsAdvice.CursorType = adOpenStatic
    rsAdvice.Open
    
    blnLimit = False
    
    Set xPath = New DOMDocument
    xPath.Load strFile
    
    '如果不包含任何元素，则退出
    If xPath.documentElement Is Nothing Then
        Set xPath = Nothing
        Screen.MousePointer = 0
        Exit Function
    End If
    
    arrSQL = Array()
    
    '读取XML内容
    Set xRoot = xPath.selectSingleNode("ClinicalPathways")
    Set xNode = xRoot.selectSingleNode("PathInfo")
    If lng路径ID = 0 Then
        '获取应用科室的情况
        strTemp1 = ""
        If Val(GetNodeValue(xNode, "通用")) = 2 Then
            Set xSubNode1 = xRoot.selectSingleNode("PathDepts")
            If Not xSubNode1 Is Nothing Then
                strSql = "Select A.ID,A.编码,A.名称" & _
                    " From 部门表 A,部门性质说明 C" & _
                    " Where A.ID=C.部门ID And C.服务对象 IN(1,3) And C.工作性质='临床'" & _
                    " Order by A.编码"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML")
                
                For Each xSubNode2 In xSubNode1.childNodes
                    rsTmp.Filter = "编码='" & GetNodeValue(xSubNode2, "编码") & "' And 名称='" & GetNodeValue(xSubNode2, "名称") & "'"
                    If Not rsTmp.EOF Then strTemp1 = strTemp1 & "," & rsTmp!ID
                Next
            
                strTemp1 = Mid(strTemp1, 2)
            End If
        End If
        
        '获取应用疾病的情况
        strValue = ""
        Set xSubNode1 = xRoot.selectSingleNode("PathDiseases")
        If Not xSubNode1 Is Nothing Then
            strTemp2 = "": strTemp3 = ""
            For Each xSubNode2 In xSubNode1.childNodes
                If Val(GetNodeValue(xSubNode2, "疾病ID")) <> 0 Then
                    strSql = "Select ID From 疾病编码目录 Where 编码=[1] And 名称=[2]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML", GetNodeValue(xSubNode2, "疾病码"), GetNodeValue(xSubNode2, "疾病名"))
                    If Not rsTmp.EOF Then strTemp2 = strTemp2 & "," & rsTmp!ID
                ElseIf Val(GetNodeValue(xSubNode2, "诊断ID")) <> 0 Then
                    strSql = "Select ID From 疾病诊断目录 Where 编码=[1] And 名称=[2]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML", GetNodeValue(xSubNode2, "诊断码"), GetNodeValue(xSubNode2, "诊断名"))
                    If Not rsTmp.EOF Then strTemp3 = strTemp3 & "," & rsTmp!ID
                End If
            Next
            If strTemp2 <> "" Or strTemp3 <> "" Then
                strValue = Mid(strTemp2, 2) & ";" & Mid(strTemp3, 2)
            End If
        End If
        
        '产生临床路径信息
        strSql = "Select ID,编码,最新版本 From 门诊路径目录 Where 分类=[1] And 名称=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML", GetNodeValue(xNode, "分类"), GetNodeValue(xNode, "名称"))
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        If Not rsTmp.EOF Then
            '新增版本或者覆盖版本
            lng路径ID = rsTmp!ID
            int版本号 = NVL(rsTmp!最新版本, 0) + 1 '可能覆盖未审核版本
            str编码 = rsTmp!编码
            arrSQL(UBound(arrSQL)) = "zl_门诊路径目录_Update(" & _
                lng路径ID & ",'" & GetNodeValue(xNode, "分类") & "','" & str编码 & "'," & _
                "'" & GetNodeValue(xNode, "名称") & "','" & GetNodeValue(xNode, "说明") & "'," & _
                Val(GetNodeValue(xNode, "适用性别")) & ",'" & GetNodeValue(xNode, "适用年龄") & "'," & _
                Val(GetNodeValue(xNode, "通用")) & "," & Val(GetNodeValue(xNode, "最大间隔时间")) & ",'" & strTemp1 & "','" & strValue & "')"
        
        Else
            '检查授权限制
            If intLimit > 0 Then
                strSql = "Select Nvl(Count(*),0) as 数量 From 门诊路径目录"
                Set rsTmp = New ADODB.Recordset
                Call zlDatabase.OpenRecordset(rsTmp, strSql, "ImportOutPathFromXML")
                If rsTmp!数量 >= intLimit Then
                    blnLimit = True
                    Set xPath = Nothing
                    Screen.MousePointer = 0
                    Exit Function
                End If
            End If
            
            '新增路径
            lng路径ID = zlDatabase.GetNextId("门诊路径目录")
            int版本号 = 1
            str编码 = GetNextCode(GetNodeValue(xNode, "分类"), 1)
            arrSQL(UBound(arrSQL)) = "zl_门诊路径目录_Insert(" & _
                "'" & GetNodeValue(xNode, "分类") & "','" & str编码 & "'," & _
                "'" & GetNodeValue(xNode, "名称") & "','" & GetNodeValue(xNode, "说明") & "'," & _
                Val(GetNodeValue(xNode, "适用性别")) & ",'" & GetNodeValue(xNode, "适用年龄") & "'," & _
                Val(GetNodeValue(xNode, "通用")) & "," & Val(GetNodeValue(xNode, "最大间隔时间")) & ",'" & strTemp1 & "','" & strValue & "'," & lng路径ID & ")"
        End If
    End If
    
    '删除版本相关的内容，重新产生
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_门诊路径版本_Delete(" & lng路径ID & "," & int版本号 & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_门诊路径版本_Update(" & lng路径ID & "," & int版本号 & "," & _
        "'" & GetNodeValue(xNode, "标准治疗时间") & "','" & GetNodeValue(xNode, "标准费用") & "'," & _
        "'" & GetNodeValue(xNode, "版本说明") & "')"
    
    '导入评估
    Set xNode = xRoot.selectSingleNode("ImportEval")
    If Not xNode Is Nothing Then
        Set xSubNode1 = xNode.selectSingleNode("Marks")
        If Not xSubNode1 Is Nothing Then
            k = 1
            Set colItemID = New Collection
            For Each xSubNode2 In xSubNode1.childNodes
                strValue = zlDatabase.GetNextId("门诊路径评估指标")
                colItemID.Add strValue, "_" & GetNodeValue(xSubNode2, "ID")
                            
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_门诊路径评估指标_Insert(" & lng路径ID & "," & int版本号 & ",NULL,1," & _
                    strValue & "," & k & ",'" & GetNodeValue(xSubNode2, "评估指标") & "'," & _
                    Val(GetNodeValue(xSubNode2, "指标类型")) & ",'" & GetNodeValue(xSubNode2, "指标结果") & "')"
                
                k = k + 1
            Next
        End If
        Set xSubNode1 = xNode.selectSingleNode("Conditions")
        If Not xSubNode1 Is Nothing Then
            For Each xSubNode2 In xSubNode1.childNodes
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_门诊路径评估条件_Insert(" & lng路径ID & "," & int版本号 & ",NULL,1," & _
                    colItemID("_" & GetNodeValue(xSubNode2, "指标ID")) & ",NULL,'" & GetNodeValue(xSubNode2, "关系式") & "'," & _
                    "'" & GetNodeValue(xSubNode2, "条件值") & "','" & GetNodeValue(xSubNode2, "条件组合") & "')"
            Next
        End If
    End If
    
    '门诊路径医嘱内容
    Set xNode = xRoot.selectSingleNode("PathAdvices")
    If Not xNode Is Nothing Then
        Set colAdviceID = New Collection
        Set colAdviceOriginalID = New Collection
        For Each xSubNode1 In xNode.childNodes
            strValue = zlDatabase.GetNextId("门诊路径医嘱内容")
            strTemp1 = GetNodeValue(xSubNode1, "ID")
            colAdviceID.Add strValue, "_" & strTemp1
            colAdviceOriginalID.Add strTemp1, "_" & strValue
        Next
        k = 1
        For Each xSubNode1 In xNode.childNodes
            blnDo = True: strTemp1 = "": strTemp2 = "": strTemp3 = ""
                
            '验证诊疗项目ID
            If Val(GetNodeValue(xSubNode1, "诊疗项目ID")) <> 0 Then
                strSql = "Select 编码,ID From 诊疗项目目录 Where 名称=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML", GetNodeValue(xSubNode1, "诊疗名称"))
                If Not rsTmp.EOF Then
                    rsTmp.Filter = "编码='" & GetNodeValue(xSubNode1, "诊疗编码") & "'"
                    If rsTmp.RecordCount > 0 Then
                        strTemp1 = rsTmp!ID
                    Else
                        rsTmp.Filter = ""
                        strTemp1 = rsTmp!ID
                    End If
                Else
                    blnDo = False
                End If
            End If
            '验证收费细目ID
            If blnDo And Val(GetNodeValue(xSubNode1, "收费细目ID")) <> 0 Then
                strSql = "Select 编码,ID From 收费项目目录 Where 名称=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML", GetNodeValue(xSubNode1, "收费名称"))
                If Not rsTmp.EOF Then
                    rsTmp.Filter = "编码='" & GetNodeValue(xSubNode1, "收费编码") & "'"
                    If rsTmp.RecordCount > 0 Then
                        strTemp2 = rsTmp!ID
                    Else
                        rsTmp.Filter = ""
                        strTemp2 = rsTmp!ID
                    End If
                Else
                    blnDo = False
                End If
            End If
            '获取导入参考
            strImportRef = IIf(Val(GetNodeValue(xSubNode1, "诊疗项目ID")) <> 0, Trim(GetNodeValue(xSubNode1, "诊疗名称")) & _
                IIf(Val(GetNodeValue(xSubNode1, "收费细目ID")) <> 0, "(" & Trim(GetNodeValue(xSubNode1, "收费名称")) & ")", ""), "" & _
                IIf(Val(GetNodeValue(xSubNode1, "收费细目ID")) <> 0, Trim(GetNodeValue(xSubNode1, "收费名称")), ""))
            '保存路径医嘱的导入状况进入临时记录集
            rsAdvice.AddNew
            rsAdvice!ID = Val(GetNodeValue(xSubNode1, "ID"))
            rsAdvice!相关id = Val(GetNodeValue(xSubNode1, "相关ID"))
            rsAdvice!导入参考 = strImportRef
            rsAdvice!导入状态 = IIf(blnDo, 1, 0)
            rsAdvice.Update
            
            If blnDo Then
                '验证执行科室ID
                If Val(GetNodeValue(xSubNode1, "执行科室ID")) <> 0 Then
                    strSql = "Select 编码,ID From 部门表 Where 名称=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML", GetNodeValue(xSubNode1, "执行科室名"))
                    If Not rsTmp.EOF Then
                        rsTmp.Filter = "编码='" & GetNodeValue(xSubNode1, "执行科室码") & "'"
                        If rsTmp.RecordCount > 0 Then
                            strTemp3 = rsTmp!ID
                        Else
                            rsTmp.Filter = ""
                            strTemp3 = rsTmp!ID
                        End If
                    End If
                End If
                
                strValue = GetNodeValue(xSubNode1, "相关ID")
                If strValue <> "" Then strValue = colAdviceID("_" & strValue)
                                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_门诊路径医嘱内容_Insert(" & _
                    colAdviceID("_" & GetNodeValue(xSubNode1, "ID")) & "," & ZVal(strValue) & "," & _
                    k & "," & Val(GetNodeValue(xSubNode1, "期效")) & "," & ZVal(strTemp1) & "," & _
                    "'" & GetNodeValue(xSubNode1, "医嘱内容") & "'," & ZVal(GetNodeValue(xSubNode1, "单次用量")) & "," & _
                    ZVal(GetNodeValue(xSubNode1, "总给予量")) & "," & ZVal(strTemp2) & "," & _
                    "'" & GetNodeValue(xSubNode1, "标本部位") & "','" & GetNodeValue(xSubNode1, "检查方法") & "'," & _
                    "'" & GetNodeValue(xSubNode1, "执行频次") & "'," & ZVal(GetNodeValue(xSubNode1, "频率次数")) & "," & _
                    ZVal(GetNodeValue(xSubNode1, "频率间隔")) & ",'" & GetNodeValue(xSubNode1, "间隔单位") & "'," & _
                    "'" & GetNodeValue(xSubNode1, "医生嘱托") & "'," & Val(GetNodeValue(xSubNode1, "执行性质")) & "," & _
                    ZVal(strTemp3) & ",'" & GetNodeValue(xSubNode1, "时间方案") & "',Null,Null," & GetNodeValue(xSubNode1, "是否缺省", 0) & "," & _
                    GetNodeValue(xSubNode1, "是否备选", 0) & ",Null," & ZVal(GetNodeValue(xSubNode1, "配方ID", 0)) & "," & ZVal(GetNodeValue(xSubNode1, "组合项目ID", 0)) & ")"
                k = k + 1
            Else
                '如果有相关ID为该医嘱的，则这些医嘱不应产生
                strValue = GetNodeValue(xSubNode1, "ID")
                For n = 0 To UBound(arrSQL)
                    If arrSQL(n) <> "" Then
                        If Split(arrSQL(n), ",")(1) = colAdviceID("_" & strValue) Then
                            '标明该医嘱不存在
                            strTemp1 = Split(Split(arrSQL(n), ",")(0), "(")(1)
                            colAdviceID.Remove "_" & colAdviceOriginalID("_" & strTemp1)
                            colAdviceID.Add "0", "_" & colAdviceOriginalID("_" & strTemp1)
                            arrSQL(n) = ""
                        End If
                    End If
                Next
                '标明该医嘱不存在
                colAdviceID.Remove "_" & strValue
                colAdviceID.Add "0", "_" & strValue
            End If
        Next
    End If
    
    '门诊路径分类
    Set xNode = xRoot.selectSingleNode("PathCategorys")
    k = 1
    For Each xSubNode1 In xNode.childNodes
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_门诊路径分类_Insert(" & lng路径ID & "," & int版本号 & "," & k & ",'" & IIf(GetNodeValue(xSubNode1, "名称") = "", xSubNode1.Text, GetNodeValue(xSubNode1, "名称")) & "')"
        k = k + 1
    Next
    
    '门诊路径阶段
    Set xNode = xRoot.selectSingleNode("PathTimeSteps")
    k = 1
    Set colStepID = New Collection
    For Each xSubNode1 In xNode.childNodes
        lng阶段ID = zlDatabase.GetNextId("门诊路径阶段")
        colStepID.Add lng阶段ID, "_" & GetNodeValue(xSubNode1, "ID")
        
        strTemp1 = GetNodeValue(xSubNode1, "父ID")
        If strTemp1 <> "" Then strTemp1 = colStepID("_" & strTemp1)
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        If strPreStep <> "" Then
            If InStr("," & strPreStep & ",", "," & GetNodeValue(xSubNode1, "ID") & ",") > 0 Then
                strtemp4 = colPreID("_" & GetNodeValue(xSubNode1, "ID"))
            End If
        End If
        arrSQL(UBound(arrSQL)) = "Zl_门诊路径阶段_Insert(" & _
            lng阶段ID & "," & lng路径ID & "," & int版本号 & "," & ZVal(strTemp1) & "," & _
            IIf(strTemp1 = "", k, GetNodeValue(xSubNode1, "序号")) & ",'" & GetNodeValue(xSubNode1, "名称") & "'," & _
            ZVal(GetNodeValue(xSubNode1, "开始天数")) & "," & ZVal(GetNodeValue(xSubNode1, "结束天数")) & "," & _
            "'" & GetNodeValue(xSubNode1, "说明") & "'," & _
            "'" & GetNodeValue(xSubNode1, "分类") & "')"
        If strTemp1 = "" Then k = k + 1
        strtemp4 = ""
        
        '阶段中的门诊路径项目
        Set xSubNode2 = xSubNode1.selectSingleNode("Items")
        If Not xSubNode2 Is Nothing Then
            Set colItemID = New Collection
            For Each xSubNode3 In xSubNode2.childNodes
                strTemp1 = "": strTemp2 = ""
                '项目关联医嘱
                lng项目ID = Val(GetNodeValue(xSubNode3, "ID"))
                Set xSubNode4 = xSubNode3.selectSingleNode("Advices")
                If Not xSubNode4 Is Nothing Then
                    For Each xSubNode5 In xSubNode4.childNodes
                        '在临时结构记录集中设置医嘱与项目的关联
                        rsAdvice.Filter = "ID=" & Val(xSubNode5.Text)
                        If rsAdvice.RecordCount <> 0 Then
                            Call rsAdvice.Update("项目ID", lng项目ID)
                        End If
                        rsAdvice.Filter = ""
                        
                        If Val(colAdviceID("_" & xSubNode5.Text)) <> 0 Then
                            strTemp1 = strTemp1 & "," & colAdviceID("_" & xSubNode5.Text)
                        End If
                    Next
                    strTemp1 = Mid(strTemp1, 2)
                End If
                
                '项目关联病历
                Set xSubNode4 = xSubNode3.selectSingleNode("EPRFiles")
                i = 1
                If Not xSubNode4 Is Nothing Then
                    For Each xSubNode5 In xSubNode4.childNodes
                        '验证病历文件ID
                        strSql = "Select ID From 病历文件列表 Where 编号=[1] And 名称=[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML", GetNodeValue(xSubNode5, "文件编号"), GetNodeValue(xSubNode5, "文件名称"))
                        If Not rsTmp.EOF Then strTemp2 = strTemp2 & ";" & rsTmp!ID & ",," & GetNodeValue(xSubNode5, "文件名称") & "," & i + 1
                    Next
                    strTemp2 = Mid(strTemp2, 2)
                End If
                
                '图标的验证：只支持固有图标
                strTemp3 = GetNodeValue(xSubNode3, "图标ID")
                If strTemp3 <> "" Then
                    If rsIcon Is Nothing Then
                        strSql = "Select ID,Nvl(性质,0) as 性质 From 临床路径图标"
                        Set rsIcon = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML")
                    End If
                    rsIcon.Filter = "ID=" & strTemp3 & " And 性质=1"
                    If rsIcon.EOF Then strTemp3 = ""
                End If
                
                strValue = zlDatabase.GetNextId("门诊路径项目")
                colItemID.Add strValue, "_" & GetNodeValue(xSubNode3, "ID")
                
                rsAdvice.Filter = "项目ID=" & lng项目ID
                
                lngCount = rsAdvice.RecordCount
                strImportRef = ""
                lng导入结果 = 1
                str组IDs = ""
                
                rsAdvice.Filter = rsAdvice.Filter & " And 导入状态=0"
                '获取导入状态
                If rsAdvice.RecordCount <> 0 Then
                    lng导入结果 = IIf(rsAdvice.RecordCount = lngCount, 0, 2)
                    '获取未导入成功医嘱的组ID
                    For n = 1 To rsAdvice.RecordCount
                        lng组ID = rsAdvice!相关id
                        If lng组ID = 0 Then lng组ID = rsAdvice!ID
                        If InStr(str组IDs & ",", "," & lng组ID & ",") = 0 Then
                            str组IDs = str组IDs & "," & lng组ID
                        End If
                        rsAdvice.MoveNext
                    Next
                End If
                If Len(str组IDs) > 0 Then str组IDs = Mid(str组IDs, 2)

                arrID = Split(str组IDs, ",")
                '获取导入参考
                For m = LBound(arrID) To UBound(arrID)
                    '过滤未导入的同一组医嘱
                    strFilter = "(项目ID = " & lng项目ID & " AND 相关ID = " & Val(arrID(m)) & ") OR (项目ID = " & lng项目ID & " AND ID=" & Val(arrID(m)) & ")"
                    rsAdvice.Filter = strFilter
                    rsAdvice.Sort = "相关ID,ID"
                    If rsAdvice.RecordCount <> 0 Then
                        For n = 1 To rsAdvice.RecordCount
                            If n = 1 And strImportRef = "" Then
                                strImportRef = rsAdvice!导入参考
                            ElseIf n = 1 And strImportRef <> "" Then
                                strImportRef = strImportRef & Chr(10) & Chr(13) & rsAdvice!导入参考 '已经有其他组医嘱已经保存在strImportRef
                            Else
                                strImportRef = strImportRef & ";" & rsAdvice!导入参考
                            End If
                            rsAdvice.MoveNext
                        Next
                    End If
                Next
   
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_门诊路径项目_Insert(" & _
                    strValue & "," & lng路径ID & "," & int版本号 & "," & lng阶段ID & "," & _
                    "'" & GetNodeValue(xSubNode3, "分类") & "'," & GetNodeValue(xSubNode3, "项目序号") & "," & _
                    "'" & GetNodeValue(xSubNode3, "项目内容") & "'," & Val(GetNodeValue(xSubNode3, "执行方式")) & _
                    ",'" & GetNodeValue(xSubNode3, "项目结果") & "'," & _
                    ZVal(strTemp3) & ",'" & strTemp1 & "','" & strTemp2 & "'," & GetNodeValue(xSubNode3, "内容要求", 0) & _
                    ",'" & Trim(strImportRef) & "'," & IIf(Trim(strImportRef) = "" And lng导入结果 = 1, "Null", lng导入结果) & ")"
            Next
        End If
        
        Set xSubNode2 = xSubNode1.selectSingleNode("StepEval")
        If Not xSubNode2 Is Nothing Then
            '评估指标
            Set xSubNode3 = xSubNode2.selectSingleNode("Marks")
            If Not xSubNode3 Is Nothing Then
                i = 1
                Set colMarkID = New Collection
                For Each xSubNode4 In xSubNode3.childNodes
                    strValue = zlDatabase.GetNextId("门诊路径评估指标")
                    colMarkID.Add strValue, "_" & GetNodeValue(xSubNode4, "ID")
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_门诊路径评估指标_Insert(" & _
                        lng路径ID & "," & int版本号 & "," & lng阶段ID & ",2," & _
                        strValue & "," & i & ",'" & GetNodeValue(xSubNode4, "评估指标") & "'," & _
                        Val(GetNodeValue(xSubNode4, "指标类型")) & ",'" & GetNodeValue(xSubNode4, "指标结果") & "')"
                    i = i + 1
                Next
            End If
            '指标条件
            Set xSubNode3 = xSubNode2.selectSingleNode("Conditions")
            If Not xSubNode3 Is Nothing Then
                For Each xSubNode4 In xSubNode3.childNodes
                    strTemp1 = GetNodeValue(xSubNode4, "指标ID")
                    If strTemp1 <> "" Then strTemp1 = colMarkID("_" & strTemp1)
                    strTemp2 = GetNodeValue(xSubNode4, "项目ID")
                    If strTemp2 <> "" Then strTemp2 = colItemID("_" & strTemp2)
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_门诊路径评估条件_Insert(" & _
                        lng路径ID & "," & int版本号 & "," & lng阶段ID & ",2," & _
                        ZVal(strTemp1) & "," & ZVal(strTemp2) & ",'" & GetNodeValue(xSubNode4, "关系式") & "'," & _
                        "'" & GetNodeValue(xSubNode4, "条件值") & "'," & Val(GetNodeValue(xSubNode4, "条件组合")) & ")"
                Next
            End If
        End If
    Next
    
    '执行提交数据
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(arrSQL)
        If CStr(arrSQL(i)) <> "" Then
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), "ImportOutPathFromXML"
        End If
    Next
    gcnOracle.CommitTrans: blnTran = False
    
    Set xPath = Nothing
    ImportOutPathFromXML = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set xPath = Nothing
End Function

Public Function CheckNotFinishPath(ByVal lng病人ID As Long, ByVal lng挂号ID As Long, ByRef lngPathID As Long, ByRef strMsg As String) As Boolean
'检查是否存在正在执行的临床路径
    Dim strSql As String, rsPati As Recordset

    On Error GoTo errH

    strSql = " Select ID, 挂号ID, 导入时间, 状态" & vbNewLine & _
             " From (Select ID, 挂号ID, 导入时间, 状态 From 病人门诊路径 Where 病人id = [1] Order By 导入时间 Desc)" & vbNewLine & _
             " Where Rownum < 2"

    Set rsPati = zlDatabase.OpenSQLRecord(strSql, "CheckNotFinishpath", lng病人ID)
    If rsPati.RecordCount > 0 Then
        If Val(NVL(rsPati!挂号ID)) = lng挂号ID Then
            CheckNotFinishPath = False                          '该病人已经有导入了临床路径
            strMsg = "该病人已经有导入了临床路径"
        ElseIf Val(NVL(rsPati!状态)) = 1 Then
            lngPathID = Val(NVL(rsPati!ID))
            CheckNotFinishPath = True                           '该病人当前存在为完成的临床路径
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckContinuePath(ByVal lngPathID As Long, ByVal lng科室ID As Long, ByVal strDiagIDs As String, ByRef strMsg As String) As Boolean
'检查是否可以继续路径
'1，这次的第一诊断（西医或者中医）必须和路径导入诊断相同；
'2，间隔时间在有效期间内
    Dim strSql As String, rsTmp As Recordset
    Dim lng阶段间隔 As Long
    
    On Error GoTo errH

    strSql = " Select Nvl(a.疾病id, a.诊断id) As 诊断id,A.科室ID" & vbNewLine & _
             " From 病人门诊路径 A, 门诊路径版本 B" & vbNewLine & _
             " Where ID = [1] And a.路径id = b.路径id And a.版本号 = b.版本号"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CheckContinuePath", lngPathID)
    If rsTmp.RecordCount > 0 Then
        lng阶段间隔 = GetIntervalTime(lngPathID)
        
        '如果诊断不同，则不能继续原有的临床路径
        If InStr("," & strDiagIDs & ",", "," & rsTmp!诊断id & ",") < 0 Then
            CheckContinuePath = False
            strMsg = "这次就诊的首要诊断和路径的导入诊断不同"
'        ElseIf lng阶段间隔 <> 0 And Val(NVL(rsTmp!最大阶段间隔)) > lng阶段间隔 Then
'            CheckContinuePath = False
'            strMsg = "超过了最大阶段间隔时间"
'        ElseIf Val(NVL(rsTmp!最大阶段间隔)) <> lng科室ID Then
'            CheckContinuePath = False
'            strMsg = "上次临床路径的科室和当前病人所在的科室不同"
        Else
            CheckContinuePath = True
        End If
        
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetIntervalTime(ByVal lngPathID As Long) As Long
'计算上次执行路径到今天为止的间隔时间
    Dim strSql As String, rsTmp As Recordset
    Dim datLatTime As Date
    
    On Error GoTo errH
    strSql = "Select Max(执行时间) as 执行时间 from 病人门诊路径执行 where 路径记录ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetIntervalTime", lngPathID)
    
    If rsTmp.RecordCount > 0 And rsTmp!执行时间 & "" <> "" Then
        datLatTime = CDate(NVL(rsTmp!执行时间))
    End If
    
    If datLatTime <> CDate(0) Then
        GetIntervalTime = DateDiff("d", datLatTime, Now)
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckOutPathSend(ByVal lng挂号ID As Long) As Boolean
'功能：检查该病人是否生成过项目
'返回：true=生成过，false=未生成过
    Dim strSql As String, rsPati As Recordset
    
    strSql = " Select Max(a.状态) As 状态" & vbNewLine & _
             " From 病人门诊路径 A, 病人门诊路径记录 B" & vbNewLine & _
             " Where a.Id = b.路径记录id And b.挂号id = [1]"
             
    On Error GoTo errH
    Set rsPati = zlDatabase.OpenSQLRecord(strSql, "CheckOutPathSend", lng挂号ID)
    If rsPati.RecordCount > 0 Then
        If Val(NVL(rsPati!状态)) <> 0 Then
            CheckOutPathSend = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetNextPhaseOut(ByVal lng阶段ID As Long) As Long
'功能：获取指定阶段的后续阶段ID
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select 父ID From 门诊路径阶段 Where id = [1] And 父ID is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "下一阶段", lng阶段ID)
    If rsTmp.RecordCount > 0 Then
        lng阶段ID = Val(rsTmp!父ID)
    End If
    
    strSql = "Select b.ID From 门诊路径阶段 a,门诊路径阶段 b " & _
            "Where a.路径ID= b.路径ID And a.版本号= b.版本号 And b.序号>a.序号 And a.ID = [1] And b.父ID Is Null And Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "下一阶段", lng阶段ID)
    
    If rsTmp.RecordCount > 0 Then GetNextPhaseOut = Val(rsTmp!ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetEPRDefineTextOut(Optional ByVal str病历IDs As String, Optional ByVal lng项目ID As Long) As String
'功能：获取路径项目对应的病历定义内容描述串
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    If lng项目ID <> 0 Then '新版电子病历和老版同时
        strSql = " Select Nvl(a.名称, b.名称) as 名称 From 门诊路径病历 A, 病历文件列表 B Where a.项目id = [2] And a.文件id = b.Id(+)" & vbNewLine & _
                 " Order by a.序号"
    ElseIf str病历IDs <> "" And lng项目ID = 0 Then '老版
        strSql = " Select /*+ Rule*/ 名称 From 病历文件列表" & _
                 " Where ID IN(Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
                 " Order by 编号"
    Else     '新版
        strSql = "select 名称 from 门诊路径病历 t where t.项目id=[2] and t.文件id is null and t.原型id IN (Select Column_Value From Table(Cast(f_Str2list([1]) As zlTools.t_Strlist))) order by 序号"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetEPRDefineTextOut", str病历IDs, lng项目ID)
    
    strSql = ""
    Do While Not rsTmp.EOF
        strSql = strSql & "、" & rsTmp!名称
        rsTmp.MoveNext
    Loop
    
    GetEPRDefineTextOut = Mid(strSql, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Check医嘱项目Out(ByVal lng执行ID As Long) As Boolean
'功能：检查指定的执行项目是否属于医嘱类
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select 1 From 病人门诊路径医嘱 Where 路径执行ID = [1] And Rownum<2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Check医嘱项目Out", lng执行ID)
    
    Check医嘱项目Out = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function CheckSameDayOfPhaseOut(ByVal lngPhase As Long, ByVal lngDay As Long) As Boolean
'功能：检查当天是否还有适用的其他后续阶段(当前阶段及分支除外)
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    '如果当前是分支阶段，则取其父ID
    strSql = "Select 父ID From 门诊路径阶段 Where ID = [1] And 父ID is Not Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取阶段", lngPhase)
    If rsTmp.RecordCount > 0 Then lngPhase = rsTmp!父ID
    
    strSql = "Select 1" & vbNewLine & _
            "From 门诊路径阶段 A, 门诊路径阶段 B" & vbNewLine & _
            "Where a.Id = [1] And a.路径id = b.路径id And a.版本号 = b.版本号 And b.序号 > a.序号" & vbNewLine & _
            "  And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取阶段", lngPhase, lngDay)
    If rsTmp.RecordCount > 0 Then CheckSameDayOfPhaseOut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiInPathOut(ByVal lng病人路径Id As Long) As Date
'功能：获取病人的进入路径的开始时间
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select a.开始时间 From 病人门诊路径 a Where a.Id =[1] "
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取入径时间", lng病人路径Id)
    If IsNull(rsTmp!开始时间) Then
        GetPatiInPathOut = zlDatabase.Currentdate
    Else
        GetPatiInPathOut = rsTmp!开始时间
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiInfoOut(lng病人ID As Long, lng挂号ID As Long) As ADODB.Recordset
    Dim strSql As String
    
    strSql = " Select Nvl(b.姓名, a.姓名) 姓名, Nvl(b.性别, a.性别) 性别, Nvl(b.年龄, a.年龄) 年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') 出生日期, b.门诊号," & vbNewLine & _
             "       b.执行状态,B.接收时间, B.完成时间, c.名称 As 科室" & vbNewLine & _
             " From 病人信息 A, 病人挂号记录 B, 部门表 C" & vbNewLine & _
             " Where a.病人id = b.病人id And b.病人id = [1] And b.Id = [2] And b.执行部门id = c.Id"
    On Error GoTo errH
    Set GetPatiInfoOut = zlDatabase.OpenSQLRecord(strSql, "读取病人数据", lng病人ID, lng挂号ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceOut(strIDs As String) As ADODB.Recordset
'功能：获取路径项目对应的医嘱记录集
    Dim strSql As String
 
    strSql = " Select /*+ rule*/ a.路径项目ID,a.医嘱内容ID,b.期效,Nvl(b.相关ID,b.ID) 相关ID,b.诊疗项目ID" & vbNewLine & _
             " From 门诊路径医嘱 A,门诊路径医嘱内容 B,(Select Column_Value As ID From Table(f_Num2list([1]))) C" & vbNewLine & _
             " Where a.医嘱内容id=b.id And a.路径项目id = c.Id" & vbNewLine & _
             " Order by b.序号"
    On Error GoTo errH
    Set GetAdviceOut = zlDatabase.OpenSQLRecord(strSql, "读取医嘱记录", strIDs)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMustDayOut(ByVal lng病人路径Id As Long, ByVal lng当前天数 As Long, Optional ByVal blnIsNotMinus As Boolean) As Long
'功能：获取病人路径执行理论上的当前天数 (=当前实际天数-曾经延迟的天数+提前天数(有可能一次提前多天))
'参数：blnIsNotMinus=是否不减去延迟时间（评估时求当前天数）
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim lng延迟天数 As Long
    Dim lng提前天数 As Long
    Dim i As Long
    Dim lng阶段实际天数 As Long
    Dim lng阶段开始天数 As Long
    Dim byt提前进度 As Byte
    
    On Error GoTo errH

    strSql = " Select Max(Decode(A.时间进度, 1, 1, 2, 2, 0)) As 阶段是否提前, C.开始天数, Nvl(C.结束天数, C.开始天数) As 结束天数," & vbNewLine & _
             "        Sum(Decode(A.时间进度, -1, 1, 0)) As 阶段延后天数, Count(1) As 阶段实际天数" & vbNewLine & _
             " From 病人门诊路径评估 A, 门诊路径阶段 C, 门诊路径阶段 D" & vbNewLine & _
             " Where a.阶段id = c.Id And c.父id = d.Id(+) And" & vbNewLine & _
             "      a.路径记录id = [1]" & vbNewLine & _
             " Group By c.开始天数, Nvl(c.结束天数, c.开始天数), a.阶段id,d.序号, c.序号" & vbNewLine & _
             " Order By Nvl(d.序号, c.序号) "

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetMustDayOut", lng病人路径Id)

    For i = 0 To rsTmp.RecordCount - 1
        '延迟天数
        lng延迟天数 = lng延迟天数 + Val(rsTmp!阶段延后天数 & "")
        '提前天数
        If Val(rsTmp!阶段是否提前 & "") = 1 Or Val(rsTmp!阶段是否提前 & "") = 2 Then
            '最后一个阶段是提前的则加1天，因为还不知道后面会选那一个阶段
            If i = rsTmp.RecordCount - 1 Or rsTmp!开始天数 & "" = rsTmp!结束天数 & "" Then
                If Val(rsTmp!阶段是否提前 & "") = 1 Then
                    lng提前天数 = lng提前天数 + 1
                ElseIf Val(rsTmp!阶段是否提前 & "") = 2 Then
                    '下一阶段提前至明天,此时不需要像“下一阶段提前到今天”再额外加一天
                End If
                rsTmp.MoveNext
            Else
                '先记录下阶段实际天数和开始天数
                lng阶段开始天数 = Val(rsTmp!开始天数 & "")
                lng阶段实际天数 = Val(rsTmp!阶段实际天数 & "")
                byt提前进度 = Val(rsTmp!阶段是否提前 & "")
                rsTmp.MoveNext
                lng提前天数 = lng提前天数 + (Val(rsTmp!开始天数 & "") - lng阶段开始天数 - lng阶段实际天数 + IIf(byt提前进度 = 2, 0, 1))
            End If
        Else
            rsTmp.MoveNext
        End If
    Next
    
    GetMustDayOut = lng当前天数 - IIf(blnIsNotMinus, 0, lng延迟天数) + lng提前天数
        
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckPathOutLogOut() As Boolean
'功能：检查是否存在病人出径登记项目
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 1 From 门诊路径报表结构 Where 报表ID = 2 And Rownum=1"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取门诊路径报表结构")
    CheckPathOutLogOut = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckDelOutPathItem(ByVal lng执行ID As Long) As Boolean
'功能：检查指定的医嘱类路径项目执行记录是否可以删除或重新生成
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strIDs As String
    Dim i As Long

    '不是当天生成的长嘱，重新生成后自动停止，不管是否发送；
    '是当天生成的长嘱，已校对但未作废，不允许取消(已停止的也不允许)，未校对的，取消时自动删除对应的医嘱。
    strSql = "Select 1 From 病人门诊路径医嘱 A, 病人门诊路径医嘱 B" & vbNewLine & _
             "Where a.路径执行id = [1] And a.病人医嘱id = b.病人医嘱id And b.路径执行id <> a.路径执行id  And rownum<2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "检查医嘱", lng执行ID)
    If rsTmp.RecordCount = 0 Then '当天生成
        strSql = "Select 1 From 病人门诊路径医嘱 B, 病人门诊医嘱记录 C Where b.路径执行id = [1] And b.病人医嘱id = c.Id And c.医嘱状态 > 1 And c.医嘱状态 <> 4 And rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "检查医嘱", lng执行ID)
        If rsTmp.RecordCount > 0 Then
            MsgBox "该项目存在已校对但未作废的医嘱，请先作废医嘱后再执行此操作。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckDelOutPathItem = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get门诊病种ID(ByVal lng病人ID As Long, ByVal lng挂号ID As Long, ByVal lngType As Long, Optional ByVal lng科室ID As Long, Optional ByRef bln中医 As Boolean = False) As ADODB.Recordset
'参数： lngType =0  按次序取门诊、中医门诊诊断;如果是中医科时, 优先级：中医门诊、门诊
'               =1  取病人除首要诊断之外的诊断，门诊诊断非第一诊断
'               =2  按次序取门诊、中医门诊诊断;如果是中医科时, 优先级：中医门诊、门诊（除开首要诊断）同时该诊断对应了首要路径。
'说明:  需排除自由录入的诊断
    Dim rsTmp As ADODB.Recordset, strSql As String

    If lngType = 0 Then                                                             '取全部诊断诊断
        bln中医 = Sys.DeptHaveProperty(lng科室ID, "中医科")
        If bln中医 Then
            strSql = " Select 疾病id,诊断id,诊断描述,诊断类型,记录来源,诊断次序" & vbNewLine & _
                     " From 病人诊断记录" & vbNewLine & _
                     " Where 记录来源 In (1,3) And 诊断类型 In (1,11) And 取消时间 Is Null And 病人id = [1] And 主页id = [2] And " & vbNewLine & _
                     "      Nvl(是否疑诊, 0) = 0 And Not (NVl(疾病ID,0)=0 and NVl(诊断ID,0)=0) " & vbNewLine & _
                     " Order By Decode(诊断类型, 11, 1, 1, 2), Decode(记录来源, 1, 4, 记录来源) Desc,诊断次序"
        Else
            strSql = " Select 疾病id,诊断id,诊断描述,诊断类型,记录来源,诊断次序" & vbNewLine & _
                     " From 病人诊断记录" & vbNewLine & _
                     " Where 记录来源 In (1,3) And 诊断类型 In (1,11) And 取消时间 Is Null And 病人id = [1] And 主页id = [2] And " & vbNewLine & _
                     "       Nvl(是否疑诊,0) = 0 And Not (NVl(疾病ID,0)=0 and NVl(诊断ID,0)=0) " & vbNewLine & _
                     " Order By 诊断类型, Decode(记录来源, 1, 4, 记录来源) Desc,诊断次序"
        End If
    ElseIf lngType = 1 Then                                                             '取非首要诊断
        strSql = " Select 疾病id, 诊断id, 诊断描述, 诊断类型, 记录来源,诊断次序" & vbNewLine & _
                 " From 病人诊断记录 " & vbNewLine & _
                 " Where 记录来源 In (1,3) And 诊断类型 In (1,11) And 取消时间 Is Null And 病人id = [1] And 主页id = [2] And 诊断次序 <> 1 Or" & vbNewLine & _
                 "      Nvl(是否疑诊, 0) = 0 And Not (NVl(疾病ID,0)=0 and NVl(诊断ID,0)=0) " & vbNewLine & _
                 " Order By 诊断类型, Decode(记录来源, 1, 4, 记录来源) Desc,诊断次序"
    ElseIf lngType = 2 Then
'        bln中医 = Sys.DeptHaveProperty(lng科室ID, "中医科")
'        If bln中医 Then
'            strSql = " Select Distinct a.Id, k.疾病id, k.诊断id, k.诊断描述, K.诊断类型, K.记录来源,k.排序 " & vbNewLine & _
'                     " From 门诊路径目录 A, 门诊路径病种 B, 门诊路径版本 C," & vbNewLine & _
'                     "     (Select Rownum As 排序, 疾病id, 诊断id, 诊断描述, 诊断类型, 记录来源 " & vbNewLine & _
'                     "       From 病人诊断记录" & vbNewLine & _
'                     "       Where 记录来源 In (1, 3) And 诊断类型 In (1,11) And 取消时间 Is Null And 病人id = [1] And 主页id = [2] And 诊断次序 <> 1 And" & vbNewLine & _
'                     "             Nvl(是否疑诊, 0) = 0 And Not (Nvl(疾病id, 0) = 0 And Nvl(诊断id, 0) = 0)" & vbNewLine & _
'                     "       Order By Decode(诊断类型,11, 3, 1, 4), Decode(记录来源, 1, 4, 记录来源) Desc, 诊断次序) K" & vbNewLine & _
'                     " Where a.Id = b.路径id And a.Id = b.路径id And a.Id = c.路径id And a.最新版本 = c.版本号 And a.性质 = 0 And b.性质 = 0 And" & vbNewLine & _
'                     "      (b.疾病id = k.疾病id Or b.诊断id = k.诊断id) And" & vbNewLine & _
'                     "      (a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 门诊路径科室 D Where a.Id = d.路径id And d.科室id = [3]))" & vbNewLine & _
'                     " Order By k.排序"
'        Else
'            strSql = " Select Distinct a.Id, k.疾病id, k.诊断id, k.诊断描述,K.诊断类型, K.记录来源,k.排序 " & vbNewLine & _
'                     " From 门诊路径目录 A, 门诊路径病种 B, 门诊路径版本 C," & vbNewLine & _
'                     "     (Select Rownum As 排序, 疾病id, 诊断id, 诊断描述, 诊断类型, 记录来源 " & vbNewLine & _
'                     "       From 病人诊断记录" & vbNewLine & _
'                     "       Where 记录来源 In (1,3) And 诊断类型 In (1,11) And 取消时间 Is Null And 病人id = [1] And 主页id = [2] And 诊断次序 <> 1 And" & vbNewLine & _
'                     "             Nvl(是否疑诊, 0) = 0 And Not (Nvl(疾病id, 0) = 0 And Nvl(诊断id, 0) = 0)" & vbNewLine & _
'                     "       Order By Sign(诊断类型 - 10), 诊断类型 Desc, Decode(记录来源, 1, 4, 记录来源) Desc, 诊断次序) K" & vbNewLine & _
'                     " Where a.Id = b.路径id And a.Id = b.路径id And a.Id = c.路径id And a.最新版本 = c.版本号 And a.性质 = 0 And b.性质 = 0 And" & vbNewLine & _
'                     "      (b.疾病id = k.疾病id Or b.诊断id = k.诊断id) And" & vbNewLine & _
'                     "      (a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 门诊路径科室 D Where a.Id = d.路径id And d.科室id = [3]))" & vbNewLine & _
'                     " Order By k.排序"
'        End If
    End If
    '记录来源:1-病历；3-首页整理
    '诊断类型:1-西医门诊诊断;11-中医门诊诊断
    '有多个诊断的情况下，根据诊断次序，只取第一个主要诊断
    '病历里面的诊断优先，主要是为了支持修正诊断。
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取病种", lng病人ID, lng挂号ID, lng科室ID)
    Set Get门诊病种ID = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiInDateOut(t_pati As TYPE_Pati) As Date
'功能：获取病人的就诊时间
'返回：就诊时间
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = " Select 执行时间 As 开始时间 " & vbNewLine & _
             "       From 病人挂号记录" & vbNewLine & _
             "       Where ID = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取就诊时间", t_pati.挂号ID)
    GetPatiInDateOut = CDate(rsTmp!开始时间)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetOutPathTable(ByVal lng疾病ID As Long, ByVal lng诊断ID As Long, ByVal lng科室ID As Long, Optional ByVal str疾病IDs As String, _
                                Optional ByVal str诊断IDs As String, Optional ByVal lng病人ID As Long, Optional ByVal lng挂号ID As Long) As ADODB.Recordset
    Dim strSql As String
    
    If str疾病IDs = "" And str诊断IDs = "" Then
        '这里加Distinct是因为，诊断id和疾病id做了绑定对应，所以查出来会有重复值
        strSql = " Select Distinct a.Id, a.分类, a.编码, a.名称, a.说明, a.适用性别, a.适用年龄, a.最新版本, c.标准治疗时间 " & vbNewLine & _
                 " From 门诊路径目录 A, 门诊路径病种 B,门诊路径版本 C" & vbNewLine & _
                 " Where a.Id = b.路径id And (b.疾病id = [1] Or b.诊断id = [2]) And a.最新版本 is not null And a.id = b.路径ID And a.最新版本 = c.版本号" & vbNewLine & _
                 " And a.Id = c.路径id And (a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 门诊路径科室 D Where a.Id = d.路径id And d.科室id = [3]))"
    Else
        strSql = " Select Distinct a.Id, a.分类, a.编码, a.名称, a.说明, a.适用性别, a.适用年龄, a.最新版本, c.标准治疗时间 " & vbNewLine & _
                 " From 门诊路径目录 A, 门诊路径病种 B,门诊路径版本 C" & vbNewLine & _
                 " Where a.Id = b.路径id And (instr(',' || [4] || ',',',' || b.疾病ID || ',')>0 " & vbNewLine & _
                 " And [4] is not null Or instr(',' || [5] || ',',',' || b.诊断ID || ',')>0 and [5] is not null) " & vbNewLine & _
                 " And a.最新版本 is not null And a.id = b.路径ID And a.最新版本 = c.版本号" & vbNewLine & _
                 " And a.Id = c.路径id And (a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 门诊路径科室 D Where a.Id = d.路径id And d.科室id = [3]))" & _
                 " And Not Exists(Select 1 From 病人门诊路径 D Where a.ID=d.路径ID And d.病人ID=[6] And D.挂号ID=[7])"
    End If
    On Error GoTo errH
    Set GetOutPathTable = zlDatabase.OpenSQLRecord(strSql, "读取路径目录", lng疾病ID, lng诊断ID, lng科室ID, str疾病IDs, str诊断IDs, lng病人ID, lng挂号ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckPatiPathOutLogOut(ByVal lng路径记录ID As Long) As Boolean
'功能：检查是否存在病人出径记录
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 1 From 病人门诊出径记录 Where 路径记录ID=[1] And Rownum=1"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取病人出径记录", lng路径记录ID)
    CheckPatiPathOutLogOut = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get阶段分类Out(Optional ByVal lng路径记录ID As Long, Optional ByVal lng阶段ID As Long) As String
'功能：获取病人使用过的阶段的分类，只有分支路径才有分类，如果使用了该分类，则病人整个路径期间只能选择该分类，所有只可能有一个分类
'参数：lng阶段ID=指定该参数时，获取指定阶段的分类
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH
    If lng阶段ID <> 0 Then
        strSql = "Select 分类 From 门诊路径阶段 Where id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取阶段分类", lng阶段ID)
    Else
        strSql = " Select a.分类" & vbNewLine & _
                 " From 门诊路径阶段 A, (Select Distinct 阶段id From 病人门诊路径执行 Where 路径记录id = [1]) B" & vbNewLine & _
                 " Where a.Id = b.阶段id And a.分类 Is Not Null And rownum<2"
    
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取阶段分类", lng路径记录ID)
    End If
    If rsTmp.RecordCount > 0 Then
        Get阶段分类Out = "" & rsTmp!分类
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPhaseNOOut(ByVal lng阶段ID As Long) As Long
'功能：获取指定阶段的序号(如果该阶段是分支，则取父阶段的序号)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select 父ID From 门诊路径阶段 Where id = [1] And 父ID is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "阶段序号", lng阶段ID)
    If rsTmp.RecordCount > 0 Then
        lng阶段ID = Val(rsTmp!父ID)
    End If
    
    strSql = "Select 序号 From 门诊路径阶段 Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "阶段序号", lng阶段ID)
    If rsTmp.RecordCount > 0 Then
        GetPhaseNOOut = Val(rsTmp!序号)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetLastPhaseNOOut(ByVal lng病人路径Id As Long, ByVal lng路径ID As Long)
'功能：获取病人指定路径最近一个阶段的序号
Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = " Select Max(Nvl(c.序号, b.序号)) 序号" & vbNewLine & _
             " From 病人门诊路径执行 A, 门诊路径阶段 B, 门诊路径阶段 C" & vbNewLine & _
             " Where a.路径记录id = [1] And a.阶段id = b.Id And b.路径id = [2] And b.父id = c.Id(+)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "阶段序号", lng病人路径Id, lng路径ID)
    
    GetLastPhaseNOOut = Val("" & rsTmp!序号)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPathDiagOut(ByVal lng挂号ID As Long, ByVal lng诊断来源 As Long, ByVal lngDiagType As Long, _
    ByVal lngDiag As Long, ByVal lng诊断ID As Long) As Boolean
'功能：检查门诊路径对应的诊断不能修改
'参数：lngDiagType：诊断类型,lngDiag=疾病ID
'返回值:F-不允许修改;T-允许修改
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select a.诊断类型, a.疾病id, a.诊断id, a.诊断来源" & vbNewLine & _
            "From 病人门诊路径 A, 病人门诊路径记录 B" & vbNewLine & _
            "Where a.Id = b.路径记录id And b.挂号id = [1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gstrSysName, lng挂号ID)
    Do While Not rsTmp.EOF
        If lngDiagType = Val(rsTmp!诊断类型 & "") And lng诊断来源 = Val(rsTmp!诊断来源 & "") And (lngDiag = Val(rsTmp!疾病id & "") Or lng诊断ID = Val(rsTmp!诊断id & "")) Then
            Exit Function
        End If
        rsTmp.MoveNext
    Loop
    CheckPathDiagOut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
