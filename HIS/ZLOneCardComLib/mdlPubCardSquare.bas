Attribute VB_Name = "mdlPubCardSquare"
Option Explicit
Public grs消费卡接口 As ADODB.Recordset

Public Function GetConsumerCardTypes(ByRef rsSquareType_Out As ADODB.Recordset, Optional ByVal blnOnlyStart As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新获取消费卡目录数据集
    '入参:
    '出参:rsSquareType_Out-返回消费卡数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-21 15:20:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long
    Dim cllData As Collection, cllTemp As Collection

    On Error GoTo errHandle
    If zlGetCardTypeRecStru(rsSquareType_Out) = False Then
        Set grs消费卡接口 = Nothing: Exit Function
    End If
        
    If zl_ExseSvr_GetConsumerCardType(cllData, blnOnlyStart) = False Then Set grs消费卡接口 = Nothing: Exit Function
    '    type_list[] C   1   支持的卡类别列表
    '        cardtype_id   N   1   id
    '        cardtype_num  N   1   编号
    '        cardtype_name C   1   名称
    '        cardtype_stname   C   1   短名
    '        prefix_text C   1   前缀文本
    '        cardno_len  N   1   卡号长度
    '        default    N   1   缺省标志
    '        fixed N   1   是否固定:1-是系统固定;0-不是系统固定
    '        strict   N   1   是否严格控制:1-是严格控制;0-不是严格控制
    '        self_make N   1   是否自制:1-是的;0-不是
    '        allow_return_cash    N   1   是否退现:1-允许;0-不允许
    '        must_all_return   N   1   是否全退:1-必需全退;0-允许部分退
    '        specpati N   1   特定病人
    '        component   C   1   部件
    '        memo    C   1   备注
    '        blnc_mode   C   1   结算方式
    '        blnc_nature N   1   结算性质
    '        pwdtxt   N   1   是否密文
    '        enabled    N   1   是否启用:1-已启用;0-未启用
    '        pwd_len N   1   密码长度
    '        pwd_len_limit   N   1   密码长度限制:0-不作限制;1-固定密码长度;-n表示密码必须输入好多个位密码以上,但不能超过密码长度
    '        pwd_rule    N   1   密码规则:０-数字和字符组成;1-仅为数字组成
    '        readcard_nature C   1   读卡性质,医疗卡读卡方式：第一位为:是否刷卡;第二位为是否扫描;第三位是否接触式读卡;第四位是否非接触式读卡。例如刷卡：'1000'
    '        keyboard_mode   N   1   键盘控制方式:：0-禁止使用软键盘;1-使用数字软键盘 ,2-使用字符软键盘
    '        def_return_cash N   1   是否缺省退现:允许退现时,默认是否退现
                
 
    For i = 1 To cllData.count
        Set cllTemp = cllData(i)
        With rsSquareType_Out
            .AddNew
                !id = cllTemp("_cardtype_id")
                !编码 = cllTemp("_cardtype_num")
                !名称 = cllTemp("_cardtype_name")
                !短名 = cllTemp("_cardtype_stname")
                
                !前缀文本 = cllTemp("_prefix_text")
                !卡号长度 = cllTemp("_cardno_len")
                !缺省标志 = cllTemp("_default")
                !是否固定 = cllTemp("_fixed")
                
                !是否严格控制 = cllTemp("_strict")
                !是否自制 = cllTemp("_self_make")
                '!是否存在帐户 = cllTemp("_exist_account")
                !是否退现 = cllTemp("_allow_return_cash")
                !是否缺省退现 = cllTemp("_def_return_cash")
                !是否全退 = cllTemp("_must_all_return")
                !部件 = cllTemp("_component")
                !备注 = cllTemp("_memo")
                
                '!特定项目 = cllTemp("_spec_item")
                !结算方式 = cllTemp("_blnc_mode")
                !结算性质 = cllTemp("_blnc_nature")
                If Val(cllTemp("_pwdtxt")) = 1 Then
                    !卡号密文 = 1
                End If
                
               ' !是否重复使用 = cllTemp("_allow_repeat_use")
                !是否启用 = cllTemp("_enabled")
                
                !密码长度 = cllTemp("_pwd_len")
                !密码长度限制 = cllTemp("_pwd_len_limit")
                !密码规则 = cllTemp("_pwd_rule")
                '!密码输入限制 = cllTemp("_pwd_require")
                '!是否模糊查找 = cllTemp("_allow_vaguefind")
                '!是否缺省密码 = cllTemp("_default_pwd")
                '!是否制卡 = cllTemp("_allow_makecard")
                '!是否发卡 = cllTemp("_allow_sendcard")
                '!是否写卡 = cllTemp("_allow_writecard")
                '!发卡控制 = cllTemp("_sendcard_sign")
                '!险类 = cllTemp("_insurance_type")
                '!发卡性质 = cllTemp("_sendcard_nature")
                '!是否转帐及代扣 = cllTemp("_transfer")
                
                !读卡性质 = cllTemp("_readcard_nature")
                !键盘控制方式 = cllTemp("_keyboard_mode")
                
                '!是否持卡消费 = cllTemp("_holding_pay")
                '!是否证件 = cllTemp("_cert_cardtype")
                '
                '!发送调用接口 = cllTemp("_advsend_buildqrcode")
                '!是否退款验卡 = cllTemp("_verfycard")
                '
                '!是否独立结算 = cllTemp("_balalone")
                '!缺省有效时间 = cllTemp("_def_valid_time")
                '!卡号识别规则 = cllTemp("_discern_rule")
                '!是否支持扫码付 = cllTemp("_scanpay")
                '!是否启用回车 = cllTemp("_enterkey_enabled")
            .Update
       End With
    Next
    If rsSquareType_Out.RecordCount <> 0 Then rsSquareType_Out.MoveFirst
    GetConsumerCardTypes = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGet消费卡接口(Optional ByVal blnOnlyStart As Boolean = True) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取消费卡接口
    '入参:blnOnlyStart-是否仅读取启用的消费卡
    '编制:刘兴洪
    '日期:2009-12-09 14:37:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
      
    If Not grs消费卡接口 Is Nothing Then
        If grs消费卡接口.State = 1 Then
            grs消费卡接口.Filter = 0
            If grs消费卡接口.RecordCount = 0 Then grs消费卡接口.MoveFirst
            Set zlGet消费卡接口 = grs消费卡接口
            Exit Function
        End If
    End If
    If GetConsumerCardTypes(grs消费卡接口, blnOnlyStart) = False Then Set grs消费卡接口 = Nothing: Exit Function
    Set zlGet消费卡接口 = grs消费卡接口
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

 

