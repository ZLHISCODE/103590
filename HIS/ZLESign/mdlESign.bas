Attribute VB_Name = "mdlESign"

Option Explicit

Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
Public Const G_STR_SPLIT As String = "&&&"

Public Enum CA_TYPE
    CA_辽宁省 = 1 '辽宁省数字证书认证中心
    CA_广西省 = 2 '广西省数字证书认证中心
    CA_重庆市 = 3 '重庆市数字证书认证中心
    CA_山东省 = 4 '山东省数字证书认证中心
    CA_吉林中心医院 = 5 '吉大正元数字证书认证中心,原名称叫 吉林中心医院 数字证书认证中心
    CA_吉林省医院 = 6 '国投安信数字证书认证中心,原名称叫 吉林省医院 数字证书认证中心
    CA_准格尔医院 = 7 '国投安信证书认证中心(内蒙),原名称叫 准格尔医院 数字证书认证中心,11年12月改成用安信的了
    CA_测试 = 8
    CA_广东 = 9 '广东数字证书认证中心(海南),还没定用不用
    CA_北京 = 10 '北京数字证书认证中心(河南)
    CA_北京CA四川 = 11 '北京数字证书认证中心(四川)
    CA_北京CA广西 = 12 '北京数字证书认证中心(广西),有时间戳
    CA_北京CA湖北 = 13 '北京数字证书认证中心(湖),有时间戳
    CA_北京CA辽宁 = 14 '北京数字证书认证中心(辽宁)
    CA_上海CA = 15     '上海数字证书认证中心 无时间戳
    CA_江苏CA = 16     '江苏数字证书认证中心  无时间戳
    CA_新疆CA = 17     '新疆数字证书认证中心  无时间戳(存在海泰key和华大key的区分)
    CA_北京CA江苏 = 18 '北京数字证书认证中心(江苏_宿迁市人民医院)
    CA_河北CA邯郸 = 19 '河北数字证书认证中心 (邯郸市第三医院)
    CA_河南CA商丘 = 20 '河南数字证书认证中心 (河南商丘市传染病医院) 20150831
    CA_辽宁三院 = 21   '辽宁省数字证书认证中心 (辽宁三院、金秋医院、中医二院)20160321
    CA_湖北 = 22       '湖北省数字证书认证管理中心有限公司
    CA_深圳 = 23       '深圳市电子商务安全证书管理有限公司
    CA_吉林安信 = 24       '吉林省安信电子认证服务有限公司
    CA_内蒙古 = 25       '内蒙古数字证书认证中心
    CA_网证通 = 26       '广东省电子商务认证有限公司(简称网证通)
End Enum

Public gintCA As CA_TYPE
Public gcnOracle As ADODB.Connection
Public gstrSysName As String
Public glngSys As Long
Public gblnShow As Boolean            '是否提示过期 False-首次提示;True-禁止提示
Public gObjFso As New FileSystemObject

Public Type TEST_SIGN
    IsInit      As Boolean      '是否已初始化
    strIniFileName As String    '配置文件名
    
    strSN As String             '序列号
    strUser As String           '用户名
    strPass As String           '密码
    strName As String           '姓名
    dateEnd As String             '证书到期时间
    
    strSignCert As String       '证书内容
    strEncCert As String        '加密证书内容
    strSignImage As String      '签名图片文件名
End Type

Public Type USER_INFO
    strID As String     '用户ID
    strName As String   '姓名
    strSignName As String  '签名
    strUserID      As String '身份证号
    lngCertID      As Long  '证书ID
    strCertSn    As String   '证书序号
    strCert   As String      '证书内容
    strCertDN As String      '证书DN
    strEncCert As String
    strCertID As String
    strPicCode As String    'BASE64签章图片编码
    strSealCode As String   '签章证书
    strTSCert As String      '时间戳证书
    strPicPath As String    '签名图片保存路径
End Type

Public Type PARA_INFO
    strSIGNIP As String      '签名服务器IP
    strSignPort As String     '签名服务器端口号
    bytSignVersion As Byte    '签名版本 RSA\SM2
    strTSIP As String        '时间戳服务器IP
    strTSPort As String       '时间戳服务器端口号
    strSignURL As String          '签名服务地址
    strTSVersion As String       '时间戳版本
    blnISTS  As Boolean       '是否启用时间戳
    blnIsSign As Boolean      '是否启用签名服务器
    blnSignPic As Boolean     '是否启用签章
    intKeyType As Integer     'Key类型 用于区别同一CA不同KEY的情况 如:安信CA
    strOption  As String      '可选参数;多个用&作为分隔符
End Type
Public Const TEST_MODE = 0      '是否测试模式运行 0- 否，1-是

Public mstrCurrPass As String     '当前密码 用于山东，记录用户输入的密码，避免用户重复输入|用于测试桩模块
Public mstrCurrUser As String     '当前用户 用于山东，记录当前用户，避免用户重复输入
Public mUserInfo As USER_INFO      '保存当前操作员信息 签名时初始化
Public gudtPara As PARA_INFO
Public gstrLogins As String           '标记已经通过登录验证的key的序列号
Public gobjComLib As Object           '公共部件对象，初始化时动态创建
Public gstrPara  As String
Public glngSign  As Long         '标记clsESign的实例数目

Public Const SWP_NOACTIVATE = &H10 '不激活窗体
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST As Long = -1
Public Const CON_TOP As Long = 262144
'准备用来使窗体始终在最前面
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter _
    As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    
Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function GetSubject(ByVal strSubject As String, ByVal strItem As String) As String
'功能：从证书主题中获取指定项目的内容
    Dim arrItem As Variant
    Dim i As Long
    
    If strSubject <> "" Then
        arrItem = Split(strSubject, ",")
        For i = 0 To UBound(arrItem)
            If UCase(arrItem(i)) Like UCase(strItem) & "=*" Then
                GetSubject = Mid(arrItem(i), InStr(arrItem(i), "=") + 1)
                Exit Function
            End If
        Next
    End If
End Function

Public Function GetSDError(ByVal lngError As Long) As String
'功能：根据山东CA接口听错误代码返回错误描述
    Dim strError As String
    
    Select Case lngError
        Case 50 'Initcontrol接口
            strError = "系统路径已经初始化" '该错误可忽略
        Case -1001 'ReadCert接口
            strError = "证书路径参数格式错误"
        Case -4100
            strError = "证书没有开通应用或者应用文件已损坏或软盘、EKEY没有插好"
        Case -4001
            strError = "证书未到应用期或系统时间设置错误"
        Case -4002
            strError = "证书没有开通相关应用"
        Case -4003
            strError = "证书的安全应用已过期或系统时间设置错误"
        Case -4004
            strError = "证书与其配置文件不匹配"
        Case 51
            strError = "证书已过期或系统时间设置错误"
        Case -9018
            strError = "没有找到证书"
        Case -9005
            strError = "证书密码错误或证书不完整"
        Case -1003
            strError = "EKEY未插好"
        Case -9003
            strError = "不能访问加密设备" '密码错误
        Case -5101
            strError = "不能实例化EKEY证书"
        Case -2001 '数字信封及获取图章接口
            strError = "验证密码错误"
        Case -3000
            strError = "加载图章文件底层库失败"
        Case -3001
            strError = "读取图章文件失败"
        Case -3002
            strError = "创建图章文件失败"
        Case -3004
            strError = "要嵌入的字符长度超过图片所容纳长度"
        Case -3006
            strError = "解压失败"
        Case -3009
            strError = "打开文件错误"
        Case -5001
            strError = "编码Base64错误"
        Case -5002
            strError = "解码Base64错误"
        Case -9009 '其他常见错误列表
            strError = "CRL不完整"
        Case -9012
            strError = "证书链不完整"
        Case -9014
            strError = "根证书无效"
        Case -9021
            strError = "私钥不存在"
        Case -9022
            strError = "算法和密钥不匹配"
        Case -9026
            strError = "证书和算法不匹配"
        Case -9027
            strError = "签名失败"
        Case -9028
            strError = "验证签名失败"
        Case -9029
            strError = "加密失败"
        Case -9030
            strError = "解密失败"
        Case -9043
            strError = "配置文件不存在"
        Case Else
            strError = "未知错误"
    End Select
    
    GetSDError = strError
End Function

Public Function PackBytes(ByVal strData As String) As Byte()
'功能：将字符串转换为字节数组
    Dim arrByte() As Byte
    Dim intAscii As Integer, intIdx As Integer
    Dim strChar As String, strHex As String
    Dim i As Integer
    
    If strData = "" Then Exit Function
    ReDim arrByte(LenB(strData) - 1)
    
    intIdx = 0
    For i = 1 To Len(strData)
        strChar = Mid(strData, i, 1)
        If strChar = Space(1) Then strChar = "+" '空格转换
        
        If strChar <> "" Then
            intAscii = Asc(strChar)
            If intAscii >= 0 Then
                arrByte(intIdx) = Asc(strChar)
                intIdx = intIdx + 1
            Else
                'Ascii<0则为中文,分高低字节转到Byte数组中
                strHex = Hex(intAscii)
                arrByte(intIdx) = Val("&H" & left(strHex, 2))
                arrByte(intIdx + 1) = Val("&H" & Right(strHex, 2))
                intIdx = intIdx + 2
            End If
        End If
    Next
    ReDim Preserve arrByte(intIdx - 1) '截掉多余的部分
    
    PackBytes = arrByte
End Function

Public Function IsUpdateRegCert(udtUser As USER_INFO, ByVal strDate As String, ByRef blnReDo As Boolean) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'功能：签名或取消签名时调用该函数（证书绑定）
'参数：udtUser-当前取值最新绑定用户信息
'      strDate-绑定证书的有效日期
'      blnRedo-True:完成注册绑定,False-未完成注册绑定
'编制:余伟节
'日期:2015-09-10
'--------------------------------------------------------------------------------------------------------------------
    Dim strTmp  As String
    Dim intDay As Integer
    Dim blnDo As Boolean
    Dim blnTip As Boolean
    Dim strMsg As String
    Dim strFileName As String
    Dim blnTrans As Boolean
    
    If mUserInfo.strCertSn <> udtUser.strCertSn Then
        '身份证号一致且姓名或者签名值相等 的情况下,允许注册
        If mUserInfo.strUserID = udtUser.strUserID And (mUserInfo.strName = udtUser.strName Or mUserInfo.strSignName = udtUser.strName) Then
            '检查旧的有效期
            If strDate <> "" Then
            '验证客户端证书有效期剩余天数
                If Not gintCA = CA_上海CA Then
                    intDay = CheckValidaty(CDate(strDate))
                Else
                    intDay = Val(strDate)
                End If
                
                If (intDay > 0) Then
                    '未过期
                    strTmp = vbCrLf & "您的证书还有" & intDay & "天过期,本机存在一个新的证书是否立即注册？" & vbCrLf
                    strTmp = strTmp & "注意:如果注册了新的,以后只能使用新的证书！"
                    If MsgBoxEx(strTmp, vbYesNo + vbInformation + vbDefaultButton2, gstrSysName) = vbYes Then
                        blnTip = True
                    Else
                        blnDo = False
                    End If
                ElseIf (intDay <= 0) Then
                    '过期
                    strTmp = "您的证书已过期 " & Abs(intDay) & " 天,本机存在一个新的证书是否立即注册？"
                    If MsgBoxEx(strTmp, vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
                        blnTip = True
                    Else
                        blnDo = False
                    End If
                End If
               
                If blnTip Then
                    strTmp = "我的信息：" & vbCrLf
                    strTmp = strTmp & Space(2) & "姓名:" & IIf(mUserInfo.strSignName = "", mUserInfo.strName, mUserInfo.strSignName) & vbTab & "身份证号：" & mUserInfo.strUserID
                    strTmp = strTmp & vbCrLf
                    strTmp = strTmp & "证书信息：" & vbCrLf
                    strTmp = strTmp & Space(2) & "使用者：" & udtUser.strName & vbTab & "绑定身份证号:" & udtUser.strUserID & vbCrLf
                    strTmp = strTmp & "核对信息无误,点""确定""完成注册。" & vbCrLf
                    strTmp = strTmp & "核对信息有误,点""取消""暂不注册。" & vbCrLf
                    If MsgBoxEx(strTmp, vbOKCancel + vbInformation + vbDefaultButton1, "注册信息") = vbOK Then
                        blnDo = True
                    Else
                        blnDo = False
                    End If
                End If
            Else
                blnDo = False   '系统注册KEY的有效结束时间为空时,暂不处理
            End If
            
            If blnDo Then
                '注册
                strTmp = "zl_人员证书记录_Insert(" & mUserInfo.strID & "," & _
                            "'" & Replace(udtUser.strCertDN, "'", "''") & "'," & _
                            "'" & Replace(udtUser.strCertSn, "'", "''") & "'," & _
                            "'" & Replace(udtUser.strCert, "'", "''") & "'," & _
                            "'" & Replace(udtUser.strEncCert, "'", "''") & "'," & _
                            "'" & Replace(udtUser.strTSCert, "'", "''") & "')"
                On Error GoTo errH
                
                '更新签名图片
                If gintCA = CA_吉林安信 Then
                    udtUser.strPicPath = ANXIN_GetSeal   '签章图片耗时,故只在更新签章图片的时候提取
                End If
                If udtUser.strPicCode <> "" Then
                    strFileName = SaveBase64ToFile("gif", udtUser.strUserID, udtUser.strPicCode)
                Else
                    strFileName = udtUser.strPicPath
                End If
                gcnOracle.BeginTrans: blnTrans = True
                If strFileName <> "" Then
                    If SaveSignPIC(Val(mUserInfo.strID), strFileName) = False Then
                        GoTo errH
                    End If
                End If
                Call gobjComLib.zlDatabase.ExecuteProcedure(strTmp, gstrSysName)
                If udtUser.strSealCode <> "" Then
                    If Not gobjComLib.Sys.Savelob(100, 14, mUserInfo.strID & "," & udtUser.strCertSn, udtUser.strSealCode, 1) Then
                        blnTrans = True
                        GoTo errH
                    End If
                End If
                gcnOracle.CommitTrans: blnTrans = False
                strMsg = "证书更新成功！"
                blnDo = False
                blnReDo = True
            Else
                blnDo = False
            End If
        Else
            strTmp = "注册信息：" & vbCrLf
            strTmp = strTmp & Space(2) & "姓名:" & IIf(mUserInfo.strSignName = "", mUserInfo.strName, mUserInfo.strSignName) & vbTab & "身份证号：" & mUserInfo.strUserID
            strTmp = strTmp & vbCrLf
            strTmp = strTmp & "证书信息：" & vbCrLf
            strTmp = strTmp & Space(2) & "使用者：" & udtUser.strName & vbTab & "绑定身份证号:" & udtUser.strUserID & vbCrLf
            strTmp = strTmp & "当前证书与注册信息不一致,不能使用!" & vbCrLf
            strMsg = strTmp
            blnDo = False
        End If
    Else
        blnDo = True '允许下一步操作（签名/取消签名）
    End If
    
    If strMsg <> "" Then
        MsgBoxEx strMsg, vbOKOnly + vbInformation, gstrSysName
    End If
    IsUpdateRegCert = blnDo
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    IsUpdateRegCert = False
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Private Function SaveSignPIC(ByVal lng人员id As Long, ByVal strFileName As String) As Boolean
    Dim rsTemp As New ADODB.Recordset, blnOk As Boolean
    
    On Error GoTo ErrHandle
    blnOk = gobjComLib.Sys.Savelob(100, 15, lng人员id, strFileName)
    SaveSignPIC = blnOk
    Exit Function
ErrHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function MsgBoxEx(prompt, Optional ByVal buttons As Long, Optional ByVal title As String) As Long
'功能:加262144 始终让Msgbox位于最顶层显示（移动护理调用时提示框未置顶显示）
    MsgBoxEx = MsgBox(prompt, CON_TOP Or buttons, title)
End Function

