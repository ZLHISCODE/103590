VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsApplicationData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public blnIsModify As Boolean               '是否被修改
Public strApplicationPageName As String         '申请单名称
Public lngApplicationPageId As Long             '申请单Id
Public strRequestTime As String             '申请时间
Public strRequestDoctor As String           '申请医生

Public lngUpdateAppNo As Long               '申请序号，只有更新时，才有申请序号
Public lngUpdateAdviceId As Long            '医嘱ID，只有更新时，才有医嘱ID
Public blnAllowUpdate As Boolean            '是否允许更新，校验后的医嘱不允许更新


Public blnIsPriority As Boolean             '是否紧急
Public blnIsAdditionalRec As Boolean        '是否补录
Public lngProjectId As Long                 '诊疗项目Id
Public lngExeType As Long                   '执行类型
Public strStartExeTime As String            '开始执行时间
Public lngExeRoomId As Long                 '执行科室
Public strExeRoomName As String             '执行科室名称
Public lngExeRoomType As Long               '执行科室类型
Public strPartMethod As String              '部位方法   格式为：部位名1;方法名1,方法名2|部位名2;方法名1,方法名2
Public strRequestAffix As String            '申请附项   格式为：附项1:附项内容|附项2:附项内容
Public strRequestAffixCfg As String         '申请附项的配置 格式为：附项1：必填，排列，要素ID
Public strDiagnoseId As String              '临床诊断Id,多个Id直接用逗号“,”分隔
Public lngRequestRoomId As Long             '申请科室Id
Public strAbstract As String                '摘要，由 gclsInsure.GetItemInfo 方法获得，提交数据时在外部赋值


Private Sub Class_Initialize()
'设置默认值
    blnAllowUpdate = True
    blnIsModify = False
End Sub

Public Function Get申请单信息(ByVal lng项目ID As Long, ByVal int应用场合 As Integer) As String
'功能：获取申请单ID申请单名称
'返回：格式 申请单ID<Split>申请单名称
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    strSQL = "Select b.Id, b.名称 From 病历单据应用 A, 病历文件列表 B Where a.诊疗项目id =[1] And a.病历文件id = b.Id And b.种类 = 7 and a.应用场合=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "clsApplicationData", lng项目ID, int应用场合)
    If Not rsTmp.EOF Then
        Get申请单信息 = rsTmp!ID & "<Split>" & rsTmp!名称
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get申请附项目配置(ByVal lng文件ID As Long) As String
'功能：获取申请单附项配置信息
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Long
    
    On Error GoTo errH
    strSQL = "select a.文件ID,a.项目,a.必填,a.排列,要素Id,b.中文名 as 要素名, a.内容,a.只读  " & _
            " from 病历单据附项 a, 诊治所见项目 b  " & _
            " where a.要素id=b.id(+) and a.文件Id=[1] order by 排列 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "clsApplicationData", lng文件ID)
    For i = 1 To rsTmp.RecordCount
        strTmp = IIF(strTmp = "", "", strTmp & "|") & rsTmp!项目 & ":" & Val(rsTmp!必填 & "") & "," & Val(rsTmp!排列 & "") & "," & NVL(rsTmp!要素ID) & ","
        rsTmp.MoveNext
    Next
    Get申请附项目配置 = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function MakePacsData(ByVal lngNo As Long, ByRef objAppPages() As clsApplicationData, Optional ByVal blnMore As Boolean) As ADODB.Recordset
'功能：通过后台查询数据组织PACS申请单对象
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim objTmp As clsApplicationData
    Dim int病人性质 As Integer '1-门诊病人，2-住院病人
    Dim lng挂号ID As Long
    Dim i As Long
    Dim lng就诊ID As Long
    Dim rsMore As ADODB.Recordset
    Dim rsOther As ADODB.Recordset
    
    On Error GoTo errH
    If lngNo = 0 Then Exit Function
    strSQL = "select a.id,a.主页ID,a.挂号单,a.开嘱科室ID,a.病人ID,b.病人性质 from 病人医嘱记录 a,病案主页 b where a.病人id=b.病人id(+) and a.主页id=b.主页id(+) and a.申请序号=[1] and a.相关ID is null order by a.序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "clsApplicationData", lngNo)
    
    With rsTmp
        If Not IsNull(!挂号单) Then
            int病人性质 = 1
            strSQL = "select a.id from 病人挂号记录 a where a.No=[1]"
            Set rsOther = zlDatabase.OpenSQLRecord(strSQL, "clsApplicationData", !挂号单 & "")
            lng就诊ID = Val(rsOther!ID & "")
        Else
            If Val(!病人性质 & "") = 1 Then
                int病人性质 = 1
            Else
                int病人性质 = 2
            End If
        End If
        
        If blnMore Then
            strSQL = "select id as 病人ID,id as 就诊ID,id as 主页ID,id as 挂号ID,id as 开嘱科室ID,id as 病人性质 from 部门表 where id=0"
            Set rsMore = zlDatabase.OpenSQLRecord(strSQL, "clsApplicationData")
            Set rsMore = zlDatabase.CopyNewRec(rsMore, True)
            rsMore.AddNew Array("病人ID", "就诊ID", "主页ID", "病人性质", "开嘱科室ID"), Array(Val(!病人ID & ""), lng就诊ID, Val(!主页ID & ""), int病人性质, Val(!开嘱科室id & ""))
        End If
        
        For i = 1 To .RecordCount
            Set objTmp = New clsApplicationData
            Set objTmp = ReadAdviceInf(Val(!ID & ""), int病人性质)
            ReDim Preserve objAppPages(i - 1)
            Set objAppPages(i - 1) = objTmp
            .MoveNext
        Next
    End With
    Set MakePacsData = rsMore
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadAdviceInf(ByVal lngAdviceID As Long, ByVal int应用场合 As Integer) As clsApplicationData
'读取当前医嘱信息
'参数：int应用场合 1-门诊，2-住院
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strCurPart As String
    Dim strCurMethod As String
    Dim strTemp As String
    Dim objReturnAppData As New clsApplicationData
    Dim str诊断内容 As String
    Dim str诊断Id As String
    
    Set objReturnAppData = Nothing
    
    '读取医嘱相关信息
    strSQL = "select a.id,a.相关Id,a.婴儿,a.诊疗项目Id,a.标本部位,a.检查方法,a.收费细目Id,a.医嘱状态,a.执行标记,a.开嘱医生,a.开嘱时间,a.开嘱科室Id," & _
            "a.执行科室ID,b.名称,a.开始执行时间,a.紧急标志,a.医嘱状态,a.申请序号 " & _
            " from 病人医嘱记录 a, 部门表 b" & _
            " where a.执行科室Id=b.Id(+) and (a.Id=[1] or a.相关Id=[1]) order by a.标本部位 desc, a.检查方法 desc"
            
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询医嘱部位方法", lngAdviceID)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    objReturnAppData.lngUpdateAdviceId = Val(NVL(rsData!ID))
    objReturnAppData.lngUpdateAppNo = Val(NVL(rsData!申请序号))
    objReturnAppData.blnAllowUpdate = Decode(Val(NVL(rsData!医嘱状态)), -1, True, 1, True, 2, True, False)
    objReturnAppData.blnIsModify = False
    
    objReturnAppData.lngProjectId = Val(NVL(rsData!诊疗项目ID))
    objReturnAppData.lngExeType = Val(NVL(rsData!执行标记))
    objReturnAppData.lngExeRoomId = Val(NVL(rsData!执行科室ID))
    objReturnAppData.strExeRoomName = NVL(rsData!名称)
    objReturnAppData.strStartExeTime = NVL(rsData!开始执行时间)
    objReturnAppData.blnIsPriority = IIF(Val(NVL(rsData!紧急标志)) = 0, False, True)
    objReturnAppData.lngRequestRoomId = Val(NVL(rsData!开嘱科室id))
    objReturnAppData.strRequestTime = NVL(rsData!开嘱时间)
    objReturnAppData.strRequestDoctor = NVL(rsData!开嘱医生)
    
    str诊断Id = GetAdviceDiag(Val(NVL(rsData!ID)), str诊断内容)
    objReturnAppData.strDiagnoseId = str诊断Id
    
    '获取部位方法
    strCurPart = ""
    strCurMethod = ""
    strTemp = ""      '部位名1;方法名1,方法名2|部位名2;方法名1,方法名2|...'<vbTab>0-常规/1-床旁/2-术中  这里不需要在后面增加执行标记
    
    While Not rsData.EOF
        If NVL(rsData!标本部位) <> strCurPart And strCurPart <> "" Then
            strTemp = strTemp & "|" & strCurPart & ";" & strCurMethod
            strCurMethod = ""
        End If
        
        If NVL(rsData!检查方法) <> "" Then
            strCurMethod = strCurMethod & IIF(strCurMethod <> "", "," & NVL(rsData!检查方法), NVL(rsData!检查方法))
        End If
        
        strCurPart = NVL(rsData!标本部位)
        
        Call rsData.MoveNext
    Wend
    
    If strCurPart <> "" Then
        strTemp = strTemp & "|" & strCurPart & ";" & strCurMethod
    End If
    
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)   '获取","之后的部位方法组合
    
    objReturnAppData.strPartMethod = strTemp
    
    '获取对应的申请单Id
    strSQL = "select a.病历文件Id, b.名称 from 病历单据应用 a, 病历文件列表 b where a.病历文件Id=b.Id and 诊疗项目Id=[1] and 应用场合=[2]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询医嘱对应申请单", objReturnAppData.lngProjectId, int应用场合)
    
    If rsData.RecordCount > 0 Then
        objReturnAppData.lngApplicationPageId = Val(NVL(rsData!病历文件ID))
        objReturnAppData.strApplicationPageName = NVL(rsData!名称) & "申请"
    End If
    
    '获取申请附项
    strSQL = "select 项目,内容 from 病人医嘱附件 where 医嘱Id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询医嘱申请附项", lngAdviceID)
    
    strTemp = ""
    If rsData.RecordCount > 0 Then
        While Not rsData.EOF
            If strTemp <> "" Then strTemp = strTemp & "|"
            
            strTemp = strTemp & NVL(rsData!项目) & ":" & NVL(rsData!内容)
            
            Call rsData.MoveNext
        Wend
    End If
    
    objReturnAppData.strRequestAffix = strTemp
    
    Set ReadAdviceInf = objReturnAppData
End Function

