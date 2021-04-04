Attribute VB_Name = "mdlPublic"
Option Explicit

Public gLastErr As String '保存最后一次错误信息
Public gbln自动读取 As Boolean '当前是否为射频卡
Public gDebug As Boolean '调试开关

'保持属性值的局部变量
Public gCol As Collection

Public gobjComLib As Object
Public gobjCommFun As Object
Public gobjDatabase As Object
Public gobjControl As Object
Public gcnConnect As ADODB.Connection

Public Property Get Cards() As Collection
    Set Cards = gCol
End Property

Private Function Add(objCard As clsCard, Optional sKey As String) As clsCard
    '创建新对象
    Dim objNewMember As clsCard
    Set objNewMember = New clsCard

    '设置传入方法的属性
    Set objNewMember = objCard
    If Len(sKey) = 0 Then
        gCol.Add objNewMember
    Else
        gCol.Add objNewMember, sKey
    End If

    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
    
End Function

Public Property Get Item(vntIndexKey As Variant) As clsCard
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = gCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    '检索集合中的元素数时使用。语法：Debug.Print x.Count
    Count = gCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    '删除集合中的元素时使用。
    'vntIndexKey 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)
    gCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    '本属性允许用 For...Each 语法枚举该集合。
    Set NewEnum = gCol.[_NewEnum]
End Property


Public Sub initCards()
    '初始化IC卡数据,固化到程序中，新增IC卡接口时，在末尾添加
    
    Dim objclsCard As clsCard
    '-- 1.Demo卡,程序测试用
    Set objclsCard = New clsCard
    objclsCard.编码 = 1
    objclsCard.接口程序名 = "zlICCard.clsICCardDev_Demo"
    objclsCard.名称 = "虚拟IC卡(测试用)"
    objclsCard.可否设置 = 1 '可以设置
    objclsCard.是否自动读取 = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & objclsCard.编码, "自动读取", 0))
    objclsCard.启用 = GetSetting("ZLSOFT", "公共模块\zlICCard", objclsCard.编码, 1) = 1
    Add objclsCard, "A" & objclsCard.编码
   '-- 2.上海医保IC卡
    Set objclsCard = New clsCard
    objclsCard.编码 = 2
    objclsCard.接口程序名 = "zl9Insure.clsInsure"
    objclsCard.名称 = "上海市医保IC卡"
    objclsCard.险类 = 413
    objclsCard.可否设置 = 0 '不能设置
    objclsCard.是否自动读取 = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & objclsCard.编码, "自动读取", 0))
    objclsCard.启用 = GetSetting("ZLSOFT", "公共模块\zlICCard", objclsCard.编码, 1) = 1
    Add objclsCard, "A" & objclsCard.编码
   '-- 3.二代证IC卡
    Set objclsCard = New clsCard
    objclsCard.编码 = 3
    objclsCard.接口程序名 = "zlICCard.clsIDcard"
    objclsCard.名称 = "第二代身份证"
    objclsCard.可否设置 = 0 '不能设置
    objclsCard.是否自动读取 = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & objclsCard.编码, "自动读取", 0))
    objclsCard.启用 = GetSetting("ZLSOFT", "公共模块\zlICCard", objclsCard.编码, 1) = 1
    Add objclsCard, "A" & objclsCard.编码
    '-- 4.明华RD系列
    Set objclsCard = New clsCard
    objclsCard.编码 = 4
    objclsCard.接口程序名 = "zlICCard.clsICCardDev_MW_RD"
    objclsCard.名称 = "明华RD系列"
    objclsCard.可否设置 = 1 '能设置
    objclsCard.是否自动读取 = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & objclsCard.编码, "自动读取", 0))
    objclsCard.启用 = GetSetting("ZLSOFT", "公共模块\zlICCard", objclsCard.编码, 1) = 1
    Add objclsCard, "A" & objclsCard.编码
    '-- 5.重庆公众城市一卡通
    Set objclsCard = New clsCard
    objclsCard.编码 = 5
    objclsCard.接口程序名 = "zlICCard.clsICCardDev_CQPubCard"
    objclsCard.名称 = "重庆公众城市一卡通"
    objclsCard.可否设置 = 0 '不需设置
    objclsCard.是否自动读取 = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & objclsCard.编码, "自动读取", 0))
    objclsCard.启用 = GetSetting("ZLSOFT", "公共模块\zlICCard", objclsCard.编码, 1) = 1
    Add objclsCard, "A" & objclsCard.编码
    
    '-- 6.诸城市人民医院射频卡接口
    Set objclsCard = New clsCard
    objclsCard.编码 = 6
    objclsCard.接口程序名 = "zlICCard.clsICCardDev_JCSRFID"
    objclsCard.名称 = "诸城市人民医院射频卡"
    objclsCard.可否设置 = 1 '不需设置
    objclsCard.是否自动读取 = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & objclsCard.编码, "自动读取", 1))
    objclsCard.启用 = GetSetting("ZLSOFT", "公共模块\zlICCard", objclsCard.编码, 1) = 1
    Add objclsCard, "A" & objclsCard.编码
    
    '-- 7.宁波一卡通
    Set objclsCard = New clsCard
    objclsCard.编码 = 7
    objclsCard.接口程序名 = "zlICCard.clsIC_NBYKT"
    objclsCard.名称 = "宁波一卡通"
    objclsCard.可否设置 = 1 '不能设置
    objclsCard.是否自动读取 = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & objclsCard.编码, "自动读取", 0))
    objclsCard.启用 = GetSetting("ZLSOFT", "公共模块\zlICCard", objclsCard.编码, 1) = 1
    Add objclsCard, "A" & objclsCard.编码

    '-- 8.剑龙D3型IC卡读写器
    '-- 2009-07-09 ZHQ 吉大口腔医院新增
    Set objclsCard = New clsCard
    objclsCard.编码 = 8
    objclsCard.接口程序名 = "zlICCard.clsICCardDev_D3IC"
    objclsCard.名称 = "剑龙D3型IC卡"
    objclsCard.可否设置 = 1 '能设置
    objclsCard.启用 = GetSetting("ZLSOFT", "公共模块\zlICCard", objclsCard.编码, 1) = 1
    Add objclsCard, "A" & objclsCard.编码
    
    '-- 9.爱生宜联射频卡
    Set objclsCard = New clsCard
    objclsCard.编码 = 9
    objclsCard.接口程序名 = "zlICCard.clsICCardDev_URF_35H"
    objclsCard.名称 = "明华URF-35H射频卡"
    objclsCard.可否设置 = 1 '不需设置
    objclsCard.是否自动读取 = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & objclsCard.编码, "自动读取", 0))
    objclsCard.启用 = GetSetting("ZLSOFT", "公共模块\zlICCard", objclsCard.编码, 1) = 1
    Add objclsCard, "A" & objclsCard.编码
    
    '-- 10.雅安一卡通
    Set objclsCard = New clsCard
    objclsCard.编码 = 10
    objclsCard.接口程序名 = "zlICCard.clsICCardDev_SLE4428"
    objclsCard.名称 = "雅安一卡通"
    objclsCard.可否设置 = 1
    objclsCard.是否自动读取 = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & objclsCard.编码, "自动读取", 0))
    objclsCard.启用 = GetSetting("ZLSOFT", "公共模块\zlICCard", objclsCard.编码, 1) = 1
    Add objclsCard, "A" & objclsCard.编码
    
    '-- 11.深圳证通金卡读写器 ZT606
    Set objclsCard = New clsCard
    objclsCard.编码 = 11
    objclsCard.接口程序名 = "zlICCard.clsICCardDev_ZT606"
    objclsCard.名称 = "深圳证通金卡读写器(ZT606)"
    objclsCard.可否设置 = 1
    objclsCard.是否自动读取 = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & objclsCard.编码, "自动读取", 0))
    objclsCard.启用 = GetSetting("ZLSOFT", "公共模块\zlICCard", objclsCard.编码, 1) = 1
    Add objclsCard, "A" & objclsCard.编码
    
    '-- 12.明华诚信 MHCX磁卡读写器  MHCX-715K(北京明华诚信科技有限公司)
    Set objclsCard = New clsCard
    objclsCard.编码 = 12
    objclsCard.接口程序名 = "zlICCard.clsICCardDev_MHCX_715K"
    objclsCard.名称 = "明华诚信MHCX磁卡读写器(MHCX_715K)"
    objclsCard.可否设置 = 1
    objclsCard.是否自动读取 = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & objclsCard.编码, "自动读取", 0))
    objclsCard.启用 = GetSetting("ZLSOFT", "公共模块\zlICCard", objclsCard.编码, 1) = 1
    Add objclsCard, "A" & objclsCard.编码
    
    '-- 13.神思四合一读卡器  SS728MQ1(山东神思电子技术股份有限公司)
    Set objclsCard = New clsCard
    objclsCard.编码 = 13
    objclsCard.接口程序名 = "zlICCard.clsICCardDev_SS728MQ1"
    objclsCard.名称 = "神思四合一读卡器(SS728MQ1)"
    objclsCard.可否设置 = 1
    objclsCard.是否自动读取 = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & objclsCard.编码, "自动读取", 0))
    objclsCard.启用 = GetSetting("ZLSOFT", "公共模块\zlICCard", objclsCard.编码, 1) = 1
    Add objclsCard, "A" & objclsCard.编码
    
End Sub

Public Sub WritLog(ByVal strDev As String, strInput As String, strOutput As String)
'    Call LogWrite("一卡通接口调试日志", 1151, "读卡接口返回", "函数名:" & strDev & ";输入:" & strInput & ";输出:" & strOutput)
End Sub

Public Function OraDataOpen(cnOracle As ADODB.Connection, ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, Optional blnMessage As Boolean = True) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strError As String
    
    On Error Resume Next
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
    End With
    If err <> 0 Then
        If blnMessage = True Then
            '保存错误信息
            strError = err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, "IC卡接口"
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, "IC卡接口"
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, "IC卡接口"
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, "IC卡接口"
            Else
                MsgBox "由于用户、口令或服务器指定错误，无法注册。", vbInformation, "IC卡接口"
            End If
        End If
        
        err.Clear
        OraDataOpen = False
        Exit Function
    End If
    OraDataOpen = True
End Function
