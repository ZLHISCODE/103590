VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsESign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private miTrusPTA As Object '用于广西New iTrusPTA
Private mCerts As Object '用于广西 New Certificates
Private mCert As Object   '用于广西 New Certificate
Private mobjTimeStamp As Object '用于广西:时间戳

Private mobjCQCA As EZCA_SignTools '用于重庆 东方中讯数字证书有限公司
Private mobjESign As Object          '用于河南商丘

Private mobjSDCA As Object '用于山东签名
Private mobjSDImg As Object '用于山东签名图章
Private mstrSDcer As String '用于山东：加密证书路径
Private mstrSDpfx As String '用于山东：签名证书路径
Private mNetONEX As Object '用于山东：时间戳
Private mstrPara As String  '用于山东:区别新老用户时间戳启用情况 1-启用,0-未启用


Private mobjJLClient As Object  '用于吉林中心医院
Private mobjJLServer As Object  '用于吉林中心医院

Private mobjJLSYY As Object '用于吉林省医院

Private mobjFile As New FileSystemObject

Private mLastCheckUSBTime As Date  ' 上次检查USB是否插入的时间，如果超过3秒，重新调用检查

Private mblnWebLogin As Boolean     '是否已经进行了WEB验证
Private mblnLNCAinit As Boolean    '辽宁签名时，如果已执行过初始化，拔掉USB再插上时，再执行初始化就过不了。
Private mblnUSETSA As Boolean       '是否使用了时间戳 －初始化成功就可以用，失败就不用

Private mTestSign As TEST_SIGN      '保存测试接口的数据

'当在.net中动态创建该对象时，不能通过方法的byref参数对数据进行返回，必须使用属性的方式进行返回
Private mlngCertId As Long          '证书ID
Private mstrTimeStamp As String     '时间戳
Private mstrTimeStampCode As String '时间戳base64编码
Private mpicSign As StdPicture      '签名图片



Private Enum ENUM_BASE64 '用于广西
    INPUT_BASE64 = 1
    CERT_ENCODE_BASE64 = 2
    OUTPUT_BASE64 = 4
End Enum

Property Get CertID() As Long
'证书ID
    CertID = mlngCertId
End Property

Property Get TimeStamp() As String
'时间戳
    TimeStamp = mstrTimeStamp
End Property


Property Get TimeStampCode() As String
'时间戳的base64编码
    TimeStampCode = mstrTimeStampCode
End Property

Property Get SignPic() As StdPicture
'签名图片
    Set SignPic = mpicSign
End Property


Public Function Initialize(cnMain As ADODB.Connection, ByVal lngSys As Long) As Boolean
'功能：初始化接口部件
'参数：cnMain=数据库连接对象
    Dim strPath As String, lngRet As Long, lngR As Long, strMsg As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    LogWrite "Initialize", "初始化签名接口,系统号" & lngSys
    Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
    Set gcnOracle = cnMain
    LogWrite "Initialize", "创建公共部件成功"
    glngSys = lngSys
    
    mLastCheckUSBTime = Now - 0.01
    
    '获取当前使用的电子签名认证中心
    
    If TEST_MODE = 0 Then
        '类似lis独立出来的系统,通过系统号和参数号不能明确参数值
        gintCA = Val(gobjComLib.zlDatabase.GetPara("电子签名认证中心", glngSys))
        LogWrite "Initialize", gintCA & "号接口"
        If gintCA = 0 Then Exit Function
    Else
        LogWrite "Initialize", "测试模式,固定为 " & CA_测试 & "号接口"
        gintCA = CA_测试
    End If
    
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    Set rsTmp = gobjComLib.zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            mstrCurrUser = Trim("" & rsTmp!User)
        End If
    End If
    LogWrite "Initialize", gintCA & "号接口开始初始化!"
    
    If gintCA = CA_辽宁省 Then
        mblnLNCAinit = False
        Err.Clear: On Error Resume Next
        Set gobjLNCA = CreateObject("LNCATOOLKITS.LNCAToolkitsCtrl.2")
        If Err.Number <> 0 Then
            MsgBoxEx Err.Number & ":" & vbCrLf & Err.Description, vbCritical, gstrSysName
        End If
        Err.Clear: On Error GoTo 0
        If gobjLNCA Is Nothing Then Exit Function
    ElseIf gintCA = CA_广西省 Then
        '可以没有安装时间戳服务器
        Set mobjTimeStamp = Nothing
        Err.Clear: On Error Resume Next
        Set miTrusPTA = CreateObject("PTA.iTrusPTA.1")
        Set mCerts = CreateObject("PTA.Certificates.1")
        Set mCert = CreateObject("PTA.Certificate.1")
        Set mobjTimeStamp = CreateObject("time.gxca")
        Err.Clear: On Error GoTo 0
    ElseIf gintCA = CA_重庆市 Then
        If mobjCQCA Is Nothing Then
            Err.Clear: On Error Resume Next
            Set mobjCQCA = New EZCA_SignTools
            If Err.Number <> 0 Then
                MsgBoxEx Err.Number & ":" & vbCrLf & Err.Description, vbCritical, gstrSysName
            End If
            Err.Clear: On Error GoTo 0
        End If
        If mobjCQCA Is Nothing Then Exit Function
        '时间戳
        mblnUSETSA = TSAWEB_initObj
    ElseIf gintCA = CA_山东省 Then
        Set mobjSDCA = Nothing
        Set mobjSDImg = Nothing
        Set mNetONEX = Nothing
        
        mstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '读取URL(IP|端口号|是否启用时间戳) 1-启用了时间戳验证
        '172.16.100.96|9198|1
        If UBound(Split(mstrPara, "|")) = 2 Then
            If Val(Split(mstrPara, "|")(2)) = 1 Then
                If mNetONEX Is Nothing Then
                    'Set mNetONEX = New NetONEX.TSAClientX
                    Err.Clear: On Error Resume Next
                    Set mNetONEX = CreateObject("NetONEX.TSAClientX.1")
                    Err.Clear: On Error GoTo 0
                    mNetONEX.ServerAddress = Split(mstrPara, "|")(0)
                    mNetONEX.ServerPort = Val(Split(mstrPara, "|")(1))
                End If
                If mNetONEX Is Nothing Then
                    MsgBoxEx "电子签名时间戳对象初始化失败！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If

        '创建山东签名控件
        Err.Clear: On Error Resume Next
        Set mobjSDCA = CreateObject("JITSECURITYTOOL.JITSecurityToolCtrl.1")
        Err.Clear: On Error GoTo 0
        If mobjSDCA Is Nothing Then Exit Function

        '创建签名图章图片处理控件
        Err.Clear: On Error Resume Next
        Set mobjSDImg = CreateObject("IMGCONVERT.ImgconvertCtrl.1")
        Err.Clear: On Error GoTo 0

        '检查是否使用文件签名
        strPath = GetSetting("ZLSOFT", "注册信息", "gstrAviPath", "") '附加文件目录
        If strPath <> "" And mobjFile.FileExists(strPath & "\server.cer") And mobjFile.FileExists(strPath & "\server.pfx") Then
            mstrSDcer = "FILE://" & Replace(strPath & "\server.cer", "\", "\\")
            mstrSDpfx = "FILE://" & Replace(strPath & "\server.pfx", "\", "\\")
        Else
            mstrSDcer = "USBCSP://.2CER"
            mstrSDpfx = "USBCSP://.2CER"
        End If

        '初始化环境:执行接口操作前必须初始化引擎
        lngRet = mobjSDCA.initcontrol()
        If Not (lngRet = 0 Or lngRet = 50) Then
            MsgBoxEx "错误" & lngRet & "：" & GetSDError(lngRet), vbExclamation, gstrSysName
            Exit Function
        End If
    ElseIf gintCA = CA_吉林中心医院 Then

        Err.Clear: On Error Resume Next
        Set mobjJLClient = Nothing
        Set mobjJLServer = Nothing

        Set mobjJLClient = CreateObject("JITComVCTK.JITVCTK.1")
        If Err.Number <> 0 Then
            MsgBoxEx "客户端签名部件 JITComVCTK.dll 没有安装。", vbExclamation, gstrSysName
            Exit Function
        End If

        Err.Clear: On Error GoTo 0
        Set mobjJLServer = CreateObject("JITClientCOMAPI.JITClientProc.1")
        If Err.Number <> 0 Then
            MsgBoxEx "服务器签名部件 JITClientCOMAPI.dll 没有安装。", vbExclamation, gstrSysName
            Exit Function
        End If

        '连接服务器
        If mobjFile.FileExists(App.Path & "\Serverinfo.ini") = False Then
            MsgBoxEx "服务器签名配置文件 Serverinfo.ini 在" & App.Path & "中没找到。", vbExclamation, gstrSysName
            Exit Function
        End If
        Dim objstr As TextStream, strIP As String, lng端口 As Long, strLine As String

        Set objstr = mobjFile.OpenTextFile(App.Path & "\Serverinfo.ini")
        Do While objstr.AtEndOfLine = False
            strLine = objstr.ReadLine
            If UCase(strLine) Like "HOST=*" Then
                strIP = Split(strLine, "=")(1)
            ElseIf UCase(strLine) Like "PORT=*" Then
                lng端口 = Val(Split(strLine, "=")(1))
            End If
        Loop
        objstr.Close
        Set objstr = Nothing

        If strIP = "" Then
            MsgBoxEx "服务器签名服务器IP未设置。", vbExclamation, gstrSysName
            Exit Function
        End If
        If lng端口 = 0 Then
            MsgBoxEx "服务器签名服务器端口未设置。", vbExclamation, gstrSysName
            Exit Function
        End If

        lngR = mobjJLServer.InitServerConnectEx(strIP, lng端口)   '根据配置文件自动确定服务器IP和端口
        If lngR <> 0 Then
            Call mobjJLServer.GetErrorMessageEx(lngR, strMsg)
            MsgBoxEx lngR & ":" & strMsg, vbExclamation, gstrSysName
            Exit Function
        End If

    ElseIf gintCA = CA_吉林省医院 Then
        Err.Clear: On Error Resume Next
        Set mobjJLSYY = Nothing
        Set mobjJLSYY = CreateObject("AXSECURITY.AXSecurityCtrl.1")
        If Err.Number <> 0 Then
            MsgBoxEx "签名控件没有安装！", vbExclamation, gstrSysName
            Exit Function
        End If
    ElseIf gintCA = CA_吉林安信 Then
        If Not ANXIN_InitObj() Then Exit Function
    ElseIf gintCA = CA_内蒙古 Then
        If Not NMG_InitObj() Then Exit Function
    ElseIf gintCA = CA_准格尔医院 Then
        If ZGRYY_initObj = False Then Exit Function
    ElseIf gintCA = CA_广东 Then
        If GDCA_initObj = False Then Exit Function
    ElseIf gintCA = CA_北京 Then
        If Not BJCA_InitObj() Then Exit Function
    ElseIf gintCA = CA_北京CA四川 Then
        If Not BJCASC_InitObj Then Exit Function
    ElseIf gintCA = CA_北京CA广西 Then
        If Not BJCAGX_InitObj Then Exit Function
    ElseIf gintCA = CA_北京CA湖北 Then
        If Not BJCAHB_InitObj Then Exit Function
    ElseIf gintCA = CA_北京CA辽宁 Then
        If Not BJCALN_InitObj Then Exit Function
    ElseIf gintCA = CA_上海CA Then
        If Not SHCA_InitObj Then Exit Function
    ElseIf gintCA = CA_江苏CA Then
        If Not JSCA_InitObj() Then Exit Function
    ElseIf gintCA = CA_新疆CA Then
        If Not XJCA_InitObj() Then Exit Function
    ElseIf gintCA = CA_北京CA江苏 Then
        If Not BJCAJS_InitObj() Then Exit Function
    ElseIf gintCA = CA_河北CA邯郸 Then
        If Not HBCA_InitObject() Then Exit Function
    ElseIf gintCA = CA_辽宁三院 Then
        If Not LNCA_Initialize() Then Exit Function
    ElseIf gintCA = CA_湖北 Then
        If Not HUBEI_InitObj() Then Exit Function
    ElseIf gintCA = CA_深圳 Then
        If Not SZCA_InitObj Then Exit Function
    ElseIf gintCA = CA_网证通 Then
        If Not WZT_InitObj Then Exit Function
    ElseIf gintCA = CA_河南CA商丘 Then
        If mobjESign Is Nothing Then
            Err.Clear: On Error Resume Next
            Set mobjESign = New clsHNCA
            If Err.Number <> 0 Then
                MsgBoxEx Err.Number & ":" & vbCrLf & Err.Description, vbCritical, gstrSysName
            End If
            Err.Clear: On Error GoTo 0
        End If
        If mobjESign Is Nothing Then Exit Function
        If Not mobjESign.HNCA_Initialize Then Exit Function
    ElseIf gintCA = CA_测试 Then
        '测试用接口。
        
        Dim objFile As TextStream
        mTestSign.strIniFileName = App.Path & "\zl9EsignTest.ini"
        If Not mobjFile.FileExists(mTestSign.strIniFileName) Then
            Set objFile = mobjFile.CreateTextFile(mTestSign.strIniFileName)
            objFile.WriteLine "SN = 001"
            objFile.WriteLine "User = ZLTEST"
            objFile.WriteLine "Pass = ZLTEST"
            objFile.WriteLine "Name = 测试帐号"
            objFile.WriteLine "EndDate = 3000-01-01"
            
            objFile.WriteLine "Cert = abcdefg"
            objFile.WriteLine "EncCert = gfedcba"
            objFile.WriteLine "SignImage = "
            objFile.Close
            Set objFile = Nothing
        End If
        
        Set objFile = mobjFile.OpenTextFile(mTestSign.strIniFileName)
        Do Until objFile.AtEndOfStream
10            strLine = objFile.ReadLine
            If UCase(strLine) Like "SN*=*" Then
20                mTestSign.strSN = Trim(Split(strLine, "=")(1))
            ElseIf UCase(strLine) Like "USER*=*" Then
30                mTestSign.strUser = Trim(Split(strLine, "=")(1))
            ElseIf UCase(strLine) Like "PASS*=*" Then
40                mTestSign.strPass = Trim(Split(strLine, "=")(1))
            ElseIf UCase(strLine) Like "NAME*=*" Then
50                mTestSign.strName = Trim(Split(strLine, "=")(1))
            ElseIf UCase(strLine) Like "ENDDATE*=*" Then
60                mTestSign.dateEnd = Trim(Split(strLine, "=")(1))
            ElseIf UCase(strLine) Like "CERT*=*" Then
70                mTestSign.strSignCert = Trim(Split(strLine, "=")(1))
            ElseIf UCase(strLine) Like "ENCCERT*=*" Then
80                mTestSign.strEncCert = Trim(Split(strLine, "=")(1))
            ElseIf UCase(strLine) Like "SIGNIMAGE*=*" Then
90                mTestSign.strSignImage = Trim(Split(strLine, "=")(1))
            End If
        Loop
        
        mTestSign.IsInit = True
    
    End If
    LogWrite "Initialize", gintCA & "号接口，初始化成功"
    
    Initialize = True
    Exit Function
errH:
    MsgBoxEx CStr(Erl()) & Err.Description
End Function

Public Function RegisterCertificate(arrCertInfo As Variant, Optional ByVal lngUserID As Long) As Boolean
    '功能：提供在HIS数据库中注册个人证书的必要信息,用于新发放或更换证书,,需要插入USB-Key
    '参数: lngUserID-人员表.ID
    '返回：arrCertInfo作为数组返回证书相关信息
    '      0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
    '      1-ClientSignCertDN:客户端签名证书主题(每人唯一)
    '      2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
    '      3-ClientSignCert:客户端签名证书内容
    '      4-ClientEncCert:客户端加密证书内容
    '      5-签名图片文件名,空串表示没有签名图片
    '      6-时间戳证书
    '      7-签章信息
        Dim strCert As String, strData As String
        Dim strPass As String, lngRet As Long

        Dim lngR As Long, strDn As String, strSN As String, strDate As String
        Dim arrDN As Variant, strUser As String, i As Integer
        Dim strBase64 As String '保存吉林省医院接口返回的图片数据(问题号34527)
        Dim strImage  As String '保存吉林省医院接口返回的图片文件路径
        Dim bSuccess As Boolean
        Dim strSQL  As String
        Dim rsTemp As ADODB.Recordset
        Dim strUserID As String
        
        On Error GoTo errH
100     If lngUserID <> 0 Then
102         If gintCA = CA_北京CA江苏 Or gintCA = CA_辽宁三院 Then
104             strSQL = "Select 身份证号,姓名 From 人员表 Where ID = [1]"
106             Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", lngUserID)
108             If Not rsTemp.EOF Then
110                 strUserID = rsTemp!身份证号 & ""
112                 If strUserID = "" Then
114                     MsgBoxEx "【" & rsTemp!姓名 & "】的身份证号为空,请联系管理员到人员管理中录入！", vbOKOnly + vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
116     arrCertInfo = Array()
118     ReDim arrCertInfo(7) '初始化返回串数组
120     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
122         arrCertInfo(i) = ""
        Next
124     If gintCA = CA_辽宁省 Then
126         With gobjLNCA
                '枚举证书：如果系统内只有1个EKey证书(既一个双证书),则不需要用户选择,否则会有提示窗口,要求用户选择数字证书
128             If .EnumClientCert = 0 Then

130                 arrCertInfo(0) = .ClientSignCertCN
132                 arrCertInfo(1) = .ClientSignCertDN
134                 arrCertInfo(2) = .ClientSignCertSN
136                 arrCertInfo(3) = .ClientSignCert
138                 arrCertInfo(4) = .ClientEncCert
                Else
140                 MsgBoxEx "获取用户数字证书失败！" & vbCrLf & vbCrLf & "错误编号:" & .errorCode & vbCrLf & "错误描述:" & .ErrorMessage, vbExclamation, gstrSysName
                    Exit Function
                End If
            End With
142     ElseIf gintCA = CA_广西省 Then
144         Set mCerts = miTrusPTA.MyCertificates
146         If mCerts.Count > 0 Then
148             Set mCert = mCerts(1) '以第1个为准

                '检查证书是否过期
150             If Int(mCert.ValidTo) <= Int(gobjComLib.zlDatabase.Currentdate) Then
152                 MsgBoxEx "用户数字证书已于 " & Format(mCert.ValidTo, "yyyy-MM-dd") & " 失效！", vbExclamation, gstrSysName
                    Exit Function
                End If
                
154             arrCertInfo(0) = mCert.CommonName
156             arrCertInfo(1) = mCert.Subject
158             arrCertInfo(2) = mCert.SerialNumber
160             arrCertInfo(3) = mCert.GetEncodedCert(CERT_ENCODE_BASE64)

162             arrCertInfo(4) = arrCertInfo(3)
            Else
164             MsgBoxEx "获取用户数字证书失败！", vbExclamation, gstrSysName
                Exit Function
            End If
166     ElseIf gintCA = CA_重庆市 Then
168         If mobjCQCA Is Nothing Then
170             MsgBoxEx "重庆CA中心的电子签名客户端环境未正确安装。", vbExclamation, gstrSysName
                Exit Function
            End If
        
            '取用户唯一标识
172         strData = mobjCQCA.WF_SelectCert()
174         If strData = "" Then
176             MsgBoxEx "证书标识获取失败。", vbCritical, gstrSysName
                Exit Function
            End If
        
178         arrCertInfo(0) = mobjCQCA.WF_GetCertOwner(strData) '姓名
180         arrCertInfo(1) = mobjCQCA.WF_GetCertDN(strData) '证书主题
182         arrCertInfo(2) = strData '存放用户唯一标识，代替证书序列号
184         arrCertInfo(3) = mobjCQCA.WF_GetPublicCert(strData) '证书公钥
186         arrCertInfo(4) = ""
188     ElseIf gintCA = CA_山东省 Then
190         If mobjSDCA Is Nothing Then
192             MsgBoxEx "山东CA中心的电子签名控件未正确安装。", vbExclamation, gstrSysName
                Exit Function
            End If

194         If CheckCurrUserPass(mstrCurrUser, strPass) = False Then Exit Function

196         lngRet = mobjSDCA.ReadCert(mstrSDcer, 2, strPass, mstrSDcer, 2, strPass)
198         If lngRet <> 0 Then
200             MsgBoxEx "错误" & lngRet & "：" & GetSDError(lngRet), vbExclamation, gstrSysName
                Exit Function
            End If
            Dim nlng As Long
202         arrCertInfo(0) = GetSubject(mobjSDCA.Subject, "CN")
204         arrCertInfo(1) = mobjSDCA.Subject
206         arrCertInfo(2) = mobjSDCA.SignCertSN
208         arrCertInfo(3) = mobjSDCA.getSignCert()
210         arrCertInfo(4) = mobjSDCA.getcryptcert()
212         strImage = App.Path & "\" & mobjSDCA.SignCertSN & "temp.gif"
214         nlng = mobjSDCA.ShowSinglePic(strPass, 0, strImage, 0)
216         If nlng = 0 Then '获取签名印章成功
218             arrCertInfo(5) = strImage '图片路径
            End If
220     ElseIf gintCA = CA_吉林中心医院 Then

222         If mobjJLClient Is Nothing Then
224             MsgBoxEx "电子签名控件未正确安装。", vbExclamation, gstrSysName
                Exit Function
            End If

226         lngR = mobjJLClient.SetCert("SC", "", "", "", "", "")
228         If lngR <> 0 Then
230             MsgBoxEx lngR & ":" & mobjJLClient.GetErrorMessage(lngR), vbExclamation, gstrSysName
                Exit Function
            End If
232         strDn = mobjJLClient.GetCertInfo("SC", 0, "") '证书主题
234         If strDn = "" Then Exit Function
236         arrDN = Split(strDn, ",")

238         For i = 0 To UBound(arrDN)
240             If Trim(arrDN(i)) Like "CN*" Then
242                 strUser = Trim(Split(arrDN(i), "=")(1))
                    Exit For
                End If
            Next
244         strSN = mobjJLClient.GetCertInfo("SC", 2, "") '序列号
246         strDate = mobjJLClient.GetCertInfo("SC", 6, "") '有效期终止日期
248         If IsDate(strDate) Then
                '检查证书是否过期
250             If Int(CDate(strDate)) <= Int(gobjComLib.zlDatabase.Currentdate) Then
252                 MsgBoxEx "用户数字证书已于 " & Format(CDate(strDate), "yyyy-MM-dd") & " 失效！", vbExclamation, gstrSysName
                    Exit Function
                End If
            End If

254         arrCertInfo(0) = strUser
256         arrCertInfo(1) = strDn
258         arrCertInfo(2) = strSN
260     ElseIf gintCA = CA_吉林省医院 Then
262         If mobjJLSYY Is Nothing Then
264             MsgBoxEx "电子签名控件未正确安装。", vbExclamation, gstrSysName
                Exit Function
            End If
266         bSuccess = mobjJLSYY.OpenCert(0, "", 1)
268         If (bSuccess) Then
270             strDn = mobjJLSYY.GetCertInfo(1, "")
272             i = InStr(1, strDn, "O=")
274             strUser = Mid(strDn, 4, i - 6)
276             strSN = mobjJLSYY.GetCertInfo(2, "")
278             strBase64 = mobjJLSYY.ReadFileFromKey(0, 2)
280             If strBase64 <> "" Then
282                 strImage = App.Path & "\" & strSN & ".gif"
284                 If Dir(strImage) <> "" Then Kill strImage
286                 If Not mobjJLSYY.B64DecodeSToFile(strBase64, strImage) Then
288                     strImage = ""
                    End If
                Else
290                 strImage = ""
                    'Exit Function
                End If
            Else
292             MsgBoxEx mobjJLSYY.GetLastError
                Exit Function
            End If

294         arrCertInfo(0) = strUser
296         arrCertInfo(1) = strDn
298         arrCertInfo(2) = strSN
            
300         arrCertInfo(5) = strImage
302     ElseIf gintCA = CA_吉林安信 Then
304         If Not ANXIN_RegCert(arrCertInfo) Then Exit Function
306     ElseIf gintCA = CA_内蒙古 Then
308         If Not NMG_RegCert(arrCertInfo) Then Exit Function
310     ElseIf gintCA = CA_准格尔医院 Then
312         If Not ZGRYY_RegCert(arrCertInfo) Then Exit Function
314     ElseIf gintCA = CA_广东 Then
316         If Not GDCA_RegCert(arrCertInfo) Then Exit Function
318     ElseIf gintCA = CA_北京 Then
320         If Not BJCA_RegCert(arrCertInfo) Then Exit Function
322     ElseIf gintCA = CA_北京CA四川 Then
324         If Not BJCASC_RegCert(arrCertInfo) Then Exit Function
326     ElseIf gintCA = CA_北京CA广西 Then
328         If Not BJCAGX_RegCert(arrCertInfo) Then Exit Function
330     ElseIf gintCA = CA_北京CA湖北 Then
332         If Not BJCAHB_RegCert(arrCertInfo) Then Exit Function
334     ElseIf gintCA = CA_北京CA辽宁 Then
336         If Not BJCALN_RegCert(arrCertInfo) Then Exit Function
338     ElseIf gintCA = CA_上海CA Then
340         If Not SHCA_RegCert(arrCertInfo) Then Exit Function
342     ElseIf gintCA = CA_江苏CA Then
344         If Not JSCA_RegCert(arrCertInfo) Then Exit Function
346     ElseIf gintCA = CA_新疆CA Then
348         If Not XJCA_RegCert(arrCertInfo) Then Exit Function
350     ElseIf gintCA = CA_北京CA江苏 Then
352         If Not BJCAJS_RegCert(arrCertInfo, strUserID) Then Exit Function
354     ElseIf gintCA = CA_河北CA邯郸 Then
356         If Not HBCA_RegCert(arrCertInfo) Then Exit Function
358     ElseIf gintCA = CA_河南CA商丘 Then
360         If Not mobjESign.HNCA_RegCert(arrCertInfo) Then Exit Function
362     ElseIf gintCA = CA_辽宁三院 Then
364         If Not LNCA_RegCert(arrCertInfo, strUserID) Then Exit Function
366     ElseIf gintCA = CA_湖北 Then
368         If Not HUBEI_RegCert(arrCertInfo) Then Exit Function
370     ElseIf gintCA = CA_深圳 Then
372         If Not SZCA_RegCert(arrCertInfo) Then Exit Function
374     ElseIf gintCA = CA_网证通 Then
376         If Not WZT_RegCert(arrCertInfo) Then Exit Function
378     ElseIf gintCA = CA_测试 Then
380         If mTestSign.IsInit Then
382              arrCertInfo(0) = mTestSign.strName
384              arrCertInfo(1) = mTestSign.strUser
386              arrCertInfo(2) = GUID()
388              arrCertInfo(3) = mTestSign.strSignCert
390              arrCertInfo(4) = mTestSign.strEncCert
392              arrCertInfo(5) = mTestSign.strSignImage
            Else
394             MsgBoxEx "测试签名控件未初始。", vbExclamation, gstrSysName
                Exit Function
            End If
        End If

396     RegisterCertificate = True
        Exit Function
errH:
398     RegisterCertificate = False
400     MsgBoxEx "第" & CStr(Erl()) & "行出现错误，" & Err.Description
End Function

Public Function ViewCertificate(ByVal lngCertID As Long) As Boolean
'功能：查看指定的在HIS数据库中注册的数字证书,不需要插入USB-Key
'参数：lngCertID=证书注册记录ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strData As String

    On Error GoTo errH

    If gintCA = CA_辽宁省 Then
        With gobjLNCA
            '1.读取证书内容
            strSQL = "Select SignCert,EncCert From 人员证书记录 Where ID=[1]"
            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", lngCertID)
            If rsTmp.EOF Then
                MsgBoxEx "未发现有效的证书注册信息。", vbInformation, gstrSysName
                Exit Function
            End If

            '2.以第三方证书进行初始化
            If .InitThirdCert(rsTmp!signCert, rsTmp!EncCert) <> 0 Then
                MsgBoxEx "数字证书初始化失败。" & vbCrLf & vbCrLf & "错误编号:" & .errorCode & vbCrLf & "错误描述:" & .ErrorMessage, vbExclamation, gstrSysName
                Exit Function
            End If

            '3.查看证书内容:4-第三方签名证书,5-第三方加密证书
            If .ViewCertWnd(4) <> 0 Then
                MsgBoxEx "无法正确查看数字证书。" & vbCrLf & vbCrLf & "错误编号:" & .errorCode & vbCrLf & "错误描述:" & .ErrorMessage, vbExclamation, gstrSysName
                Exit Function
            End If
        End With
    ElseIf gintCA = CA_广西省 Then
        '需要插入USB-Key
        MsgBoxEx "将查看密钥盘上的证书内容，按确定继续。", vbExclamation, gstrSysName

        Set mCerts = miTrusPTA.MyCertificates
        If mCerts.Count > 0 Then
            Set mCert = mCerts(1) '以第1个为准
            mCert.View '密钥盘上的，非当前选择的
        Else
            MsgBoxEx "没有发现有效的数字证书。", vbExclamation, gstrSysName
            Exit Function
        End If
    ElseIf gintCA = CA_北京CA广西 Then
       Call ShowCert(lngCertID)
    ElseIf gintCA = CA_重庆市 Then
        If mobjCQCA Is Nothing Then
            MsgBoxEx "重庆CA中心的电子签名客户端环境未正确安装。", vbExclamation, gstrSysName
            Exit Function
        End If

        '读取证书内容
        strSQL = "Select CertSN From 人员证书记录 Where ID=[1]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", lngCertID)
        If rsTmp.EOF Then
            MsgBoxEx "未发现有效的证书注册信息。", vbInformation, gstrSysName
            Exit Function
        End If

        '以用户标识查看证书(存在序列号字段中)
        Call mobjCQCA.WF_ShowCert(rsTmp!CertSN)
    ElseIf gintCA = CA_山东省 Then
        If mobjSDCA Is Nothing Then
            MsgBoxEx "山东CA中心的电子签名控件未正确安装。", vbExclamation, gstrSysName
            Exit Function
        End If

        '读取证书内容
        strSQL = "Select CertDN,CertSN,注册时间 From 人员证书记录 Where ID=[1]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", lngCertID)
        If rsTmp.EOF Then
            MsgBoxEx "未发现有效的证书注册信息。", vbInformation, gstrSysName
            Exit Function
        End If

        MsgBoxEx "证书信息" & vbCrLf & _
            vbCrLf & "颁发给：" & GetSubject(rsTmp!CertDn, "CN") & _
            vbCrLf & "证书主题：" & rsTmp!CertDn & _
            vbCrLf & "证书序列号：" & rsTmp!CertSN & _
            vbCrLf & "注册时间：" & Format(rsTmp!注册时间, "yyyy-MM-dd HH:mm:ss"), vbInformation, gstrSysName
    ElseIf gintCA = CA_吉林中心医院 Then
        If mobjJLClient Is Nothing Then
            MsgBoxEx "客户端签名部件 JITComVCTK.dll 没有安装。", vbExclamation, gstrSysName
            Exit Function
        End If
        Call ShowCert(lngCertID)
        '读取证书内容
    ElseIf gintCA = CA_吉林省医院 Then
        If mobjJLSYY Is Nothing Then
            MsgBoxEx "签名部件没有安装。", vbExclamation, gstrSysName
            Exit Function
        End If
        Call ShowCert(lngCertID)
    ElseIf gintCA = CA_准格尔医院 Then
        Call ShowCert(lngCertID)
    ElseIf gintCA = CA_广东 Then
        Call ShowCert(lngCertID)
    ElseIf gintCA = CA_江苏CA Then
        Call ShowCert(lngCertID)
    ElseIf gintCA = CA_北京CA辽宁 Then
       '读取证书内容
        strSQL = "Select SignCert,EncCert From 人员证书记录 Where ID=[1]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", lngCertID)
        If rsTmp.EOF Then
            MsgBoxEx "未发现有效的证书注册信息。", vbInformation, gstrSysName
            Exit Function
        Else
            MsgBoxEx "接口暂未提供此项功能！"
        End If
    ElseIf gintCA = CA_新疆CA Then
        Call ShowCert(lngCertID)
    ElseIf gintCA = CA_北京CA江苏 Or gintCA = CA_北京CA湖北 Or gintCA = CA_吉林安信 Or CA_内蒙古 Then
        Call ShowCert(lngCertID)
    ElseIf gintCA = CA_河北CA邯郸 Or gintCA = CA_网证通 Then
        Call ShowCert(lngCertID)
    ElseIf gintCA = CA_河南CA商丘 Then
        Call ShowCert(lngCertID)
    ElseIf gintCA = CA_辽宁三院 Then
        Call ShowCert(lngCertID)
    ElseIf gintCA = CA_湖北 Then
        Call ShowCert(lngCertID)
    ElseIf gintCA = CA_深圳 Then
        Call ShowCert(lngCertID)
    ElseIf gintCA = CA_测试 Then
        '读取证书内容
        strSQL = "Select CertDN,CertSN,注册时间 From 人员证书记录 Where ID=[1]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", lngCertID)
        If rsTmp.EOF Then
            MsgBoxEx "未发现有效的证书注册信息。", vbInformation, gstrSysName
            Exit Function
        End If
        MsgBoxEx "证书信息" & vbCrLf & _
            vbCrLf & "颁发给：" & mTestSign.strName & _
            vbCrLf & "证书主题：" & mTestSign.strUser & _
            vbCrLf & "证书序列号：" & mTestSign.strSN & _
            vbCrLf & "注册时间：" & Format(rsTmp!注册时间, "yyyy-MM-dd HH:mm:ss"), vbInformation, gstrSysName
    End If

    ViewCertificate = True
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Sub ShowCert(ByVal lngCertID As Long)
        Dim strSQL As String, rsTmp As ADODB.Recordset
        Dim strName As String
        Dim strCertUserID As String
        
        On Error GoTo errH
100     strSQL = "Select CertDN,CertSN,注册时间 From 人员证书记录 Where ID=[1]"
102     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", lngCertID)
104     If rsTmp.EOF Then
106         MsgBoxEx "未发现有效的证书注册信息。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Select Case gintCA
        Case CA_江苏CA
            Call JSCA_GetCertList(strName)
        Case CA_北京CA江苏
            Call BJCAJS_GetCertList(strName, , , strCertUserID)
        Case Else
            strName = rsTmp!CertDn & ""
        End Select
        
108     MsgBoxEx "证书信息" & vbCrLf & _
            vbCrLf & "颁发给：" & strName & IIf(strCertUserID <> "", "|" & strCertUserID, "") & _
            vbCrLf & "证书主题：" & rsTmp!CertDn & _
            vbCrLf & "证书序列号：" & rsTmp!CertSN & _
            vbCrLf & "注册时间：" & Format(rsTmp!注册时间, "yyyy-MM-dd HH:mm:ss"), vbInformation, gstrSysName
        Exit Sub
errH:
110     MsgBoxEx "ShowCert 错误出现在：" & CStr(Erl()) & "行," & Err.Description, vbExclamation, gstrSysName
End Sub

Public Function CheckCertificate(ByVal strDBUser As String) As Boolean
    '功能：检查当前USB-Key是否正确,需要插入USB-Key。
    '说明：用于某些功能(如取消签名)时检查Key盘是否插入。
    '参数：strDBUser=数据库用户名,用于取证书序列号来初始化证书
    '返回：True-注册|注册且停用；False-未注册；
    '      注册且未停用的情况下，检查是否插入KEY 正常返回True
    
        Dim rsTmp As New ADODB.Recordset, rsTemp As New ADODB.Recordset
        Dim strSQL  As String, lngRet As Long
        Dim strData As String, strPass As String
        Dim bSuccess As Boolean
        Dim lngR As Long, strDn As String, strSN As String, strDate As String
        Dim strMsg As String
        Dim lngTmp As Long
        Dim blnReDo As Boolean, blnRet As Boolean
        
        On Error GoTo errH

        '读取证书序列号
100     If strDBUser <> "" Then
102         strSQL = "Select B.ID,B.CertSN,B.CertDN,B.SignCert,B.时间戳证书 From 上机人员表 A,人员证书记录 B " & _
                " Where A.人员ID=B.人员ID And A.用户名=[1] Order by B.注册时间 Desc,B.ID Desc"
        Else
104         strSQL = "Select B.ID,B.CertSN,B.CertDN,B.SignCert,B.时间戳证书 From 上机人员表 A,人员证书记录 B " & _
                " Where A.人员ID=B.人员ID And A.用户名=USER Order by B.注册时间 Desc,B.ID Desc"
        End If
106     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", strDBUser)

108     If rsTmp.EOF Then
110         MsgBoxEx "你的数字证书还没有注册。", vbInformation, gstrSysName
            Exit Function
        End If
        strSQL = "select b.姓名,b.ID as 人员ID,b.身份证号,b.签名 from 上机人员表 a,人员表 b where a.人员id=b.id and a.用户名=[1]"
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", strDBUser)
        
         '--获取当前操作员信息----------------------------------------
        mUserInfo.strID = Trim(rsTemp!人员ID & "")
        mUserInfo.strName = Trim(rsTemp!姓名 & "")
        mUserInfo.strSignName = Trim(rsTemp!签名 & "")
        mUserInfo.strUserID = Trim(rsTemp!身份证号 & "")
        mUserInfo.strCertSN = Trim("" & rsTmp!CertSN)
        mUserInfo.strCert = Trim("" & rsTmp!signCert)
        mUserInfo.strTSCert = Trim("" & rsTmp!时间戳证书)
        mUserInfo.strCertDN = Trim("" & rsTmp!CertDn)
        '-------------------------------------------
        If Not rsTemp.EOF Then
            If CertificateStoped(rsTemp!姓名) Then
                CheckCertificate = True  '注册且停用
                Exit Function
            End If
        End If
112     If gintCA = CA_辽宁省 Then
114         With gobjLNCA
116             If Not mblnLNCAinit Or mstrCurrUser <> strDBUser Then
118                 If .InitClientCertByCertSN("" & rsTmp!CertSN) <> 0 Then
120                     MsgBoxEx "数字证书初始化失败。" & vbCrLf & vbCrLf & "错误编号:" & .errorCode & vbCrLf & "错误描述:" & .ErrorMessage, vbExclamation, gstrSysName
                        Exit Function
                    End If
122                 mblnLNCAinit = True
                End If
            End With
124     ElseIf gintCA = CA_广西省 Then
126         Set mCerts = miTrusPTA.MyCertificates
128         If mCerts.Count > 0 Then
130             Set mCert = mCerts(1) '以第1个为准
132             If mCert.SerialNumber <> rsTmp!CertSN Then
134                 MsgBoxEx "当前数字证书与注册的不相符。", vbExclamation, gstrSysName
                    Exit Function
                End If
            Else
136             MsgBoxEx "没有发现有效的数字证书。", vbExclamation, gstrSysName
                Exit Function
            End If
138     ElseIf gintCA = CA_重庆市 Then
140         If mobjCQCA Is Nothing Then
142             MsgBoxEx "重庆CA中心的电子签名客户端环境未正确安装。", vbExclamation, gstrSysName
                Exit Function
            End If
        
            '取用户唯一标识
144         strData = mobjCQCA.WF_SelectCert()
146         If strData = "" Then
148             MsgBoxEx "证书标识获取失败。", vbCritical, gstrSysName
                Exit Function
            End If
150     ElseIf gintCA = CA_山东省 Then
152         If mobjSDCA Is Nothing Then
154             MsgBoxEx "山东CA中心的电子签名控件未正确安装。", vbExclamation, gstrSysName
                Exit Function
            End If

156         If CheckCurrUserPass(strDBUser, strPass) = False Then Exit Function

158         lngRet = mobjSDCA.ReadCert(mstrSDcer, 2, strPass, mstrSDcer, 2, strPass)
160         If lngRet <> 0 Then
162             MsgBoxEx "错误" & lngRet & "：" & GetSDError(lngRet), vbExclamation, gstrSysName
                Exit Function
            End If

164         If mobjSDCA.SignCertSN <> rsTmp!CertSN Then
166             MsgBoxEx "当前数字证书与注册的不相符。", vbExclamation, gstrSysName
                Exit Function
            End If
168     ElseIf gintCA = CA_吉林中心医院 Then
            '设置证书
170         If mobjJLClient Is Nothing Then
172             MsgBoxEx "客户端签名部件 JITComVCTK.dll 没有安装。", vbExclamation, gstrSysName
                Exit Function
            End If

174         lngR = mobjJLClient.SetCert("SC", "", "", "", "", "")
176         strDn = mobjJLClient.GetCertInfo("SC", 0, "") '证书主题
178         strSN = mobjJLClient.GetCertInfo("SC", 2, "") '序列号

180         If strSN <> rsTmp!CertSN Then
182             MsgBoxEx "当前数字证书与注册的不相符。", vbExclamation, gstrSysName
                Exit Function
            End If

184         strDate = mobjJLClient.GetCertInfo("SC", 6, "") '有效期终止日期
186         If IsDate(strDate) Then
                '检查证书是否过期
188             If Int(CDate(strDate)) <= Int(gobjComLib.zlDatabase.Currentdate) Then
190                 MsgBoxEx "用户数字证书已于 " & Format(CDate(strDate), "yyyy-MM-dd") & " 失效！", vbExclamation, gstrSysName
                    Exit Function
                End If
            End If

            '设置为使用 当前证书
192         lngR = mobjJLClient.SetCert("SC", strDn, strSN, "", "", "")
194         If lngR <> 0 Then
196             MsgBoxEx lngR & ":" & mobjJLClient.GetErrorMessage(lngR), vbExclamation, gstrSysName
                Exit Function
            End If
198         mLastCheckUSBTime = Now
200     ElseIf gintCA = CA_吉林省医院 Then
202         If mobjJLSYY Is Nothing Then
204             MsgBoxEx "电子签名控件未正确安装。", vbExclamation, gstrSysName
                Exit Function
            End If
206         bSuccess = mobjJLSYY.SetSignerCert(2, "" & rsTmp!CertSN)
208         If Not bSuccess Then
210             MsgBoxEx "没有找到指定的证书!" & vbNewLine & rsTmp!CertSN, vbExclamation, gstrSysName
                Exit Function
            End If
212         mLastCheckUSBTime = Now
214     ElseIf gintCA = CA_准格尔医院 Then
216         strSN = Trim$("" & rsTmp!CertSN)
            If Not ZGRYY_CheckCert(strSN) Then Exit Function
        ElseIf gintCA = CA_吉林安信 Then
            blnRet = ANXIN_CheckCert(blnReDo)
            If Not blnRet And Not blnReDo Then Exit Function
        ElseIf gintCA = CA_内蒙古 Then
            blnRet = NMG_CheckCert(blnReDo)
            If Not blnRet And Not blnReDo Then Exit Function
222     ElseIf gintCA = CA_广东 Then
224         strSN = Trim("" & rsTmp!CertSN)
226         If Not GDCA_CheckCert(strSN) Then
228             MsgBoxEx "当前数字证书与注册的不相符。", vbExclamation, gstrSysName
                Exit Function
            End If
230     ElseIf gintCA = CA_北京 Then
232         strSN = Trim$("" & rsTmp!CertSN)
234         blnRet = BJCA_CheckCert(blnReDo)
            If Not blnRet And Not blnReDo Then Exit Function
236     ElseIf gintCA = CA_北京CA四川 Then
238         strSN = Trim$("" & rsTmp!CertSN)
240         If Not BJCASC_CheckCert(strSN) Then Exit Function
        ElseIf gintCA = CA_北京CA广西 Then
            blnRet = BJCAGX_CheckCert(blnReDo)
            If Not blnRet And Not blnReDo Then Exit Function
        ElseIf gintCA = CA_北京CA湖北 Then
            If Not BJCAHB_CheckCert(Trim("" & rsTmp!CertSN)) Then Exit Function
        ElseIf gintCA = CA_北京CA辽宁 Then
            strSN = Trim$("" & rsTmp!CertSN)
            If Not BJCALN_CheckCert(strSN) Then Exit Function
        ElseIf gintCA = CA_江苏CA Then
            If Not JSCA_CheckCert(Trim$("" & rsTmp!CertSN)) Then Exit Function
        ElseIf gintCA = CA_新疆CA Then
            If Not XJCA_CheckCert(Trim$("" & rsTmp!CertSN)) Then Exit Function
        ElseIf gintCA = CA_北京CA江苏 Then
            blnRet = BJCAJS_CheckCert(blnReDo)
            If Not blnRet And Not blnReDo Then Exit Function
        ElseIf gintCA = CA_河北CA邯郸 Then
            blnRet = HBCA_CheckCert(blnReDo)
            If Not blnRet And Not blnReDo Then Exit Function
        ElseIf gintCA = CA_河南CA商丘 Then
            blnRet = mobjESign.HNCA_CheckCert(blnReDo)
            If Not blnRet And Not blnReDo Then Exit Function
        ElseIf gintCA = CA_辽宁三院 Then
            blnRet = LNCA_CheckCert(blnReDo)
            If Not blnRet And Not blnReDo Then Exit Function
        ElseIf gintCA = CA_湖北 Then
            If Not HUBEI_CheckCert(Trim("" & rsTmp!CertSN)) Then Exit Function
        ElseIf gintCA = CA_深圳 Then
            If Not SZCA_CheckCert(Trim("" & rsTmp!CertSN)) Then Exit Function
        ElseIf gintCA = CA_网证通 Then
            If Not WZT_CheckCert() Then Exit Function
        ElseIf gintCA = CA_上海CA Then
            blnRet = SHCA_CheckCert(Trim("" & rsTmp!CertSN))
            If Not blnRet And Not blnReDo Then Exit Function
242     ElseIf gintCA = CA_测试 Then
244         If Not mTestSign.IsInit Then
246             MsgBoxEx "测试CA中心未正确初始化。", vbExclamation, gstrSysName
                Exit Function
            End If
            If Not mobjFile.FileExists(mTestSign.strIniFileName) Then
                MsgBoxEx "没有发现有效的数字证书。", vbExclamation, gstrSysName
                Exit Function
            End If
248         strMsg = ""
            '要求用户输入密码，输过了就不用再输入
            
250         If CheckCurrUserPass(strDBUser, strPass) = False Then Exit Function
            
        
            '检查传入的用户，注册的证书号是否与当前证书号一致
252         'If rsTmp!CertSN <> mTestSign.strSN Then strMsg = strMsg & IIf(strMsg = "", "", vbNewLine) & "当前数字证书与注册的不相符!"
            
            '检查密码
254         If strPass <> mTestSign.strPass Then
                strMsg = strMsg & IIf(strMsg = "", "", vbNewLine) & "密码错误！"
                mstrCurrPass = ""
            End If
            '检验效期
256         If IsDate(mTestSign.dateEnd) Then
                lngTmp = DateDiff("D", gobjComLib.zlDatabase.Currentdate, CDate(mTestSign.dateEnd))
                If lngTmp > 0 And lngTmp <= 30 And Not gblnShow Then
                    MsgBoxEx "您的证书还有" & lngTmp & "天过期！", vbExclamation, gstrSysName
                    gblnShow = True
258             ElseIf Int(CDate(mTestSign.dateEnd)) <= Int(gobjComLib.zlDatabase.Currentdate) Then
                    strMsg = strMsg & IIf(strMsg = "", "", vbNewLine) & "证书已过期！"
                End If
            End If
260         If strMsg <> "" Then
262             MsgBoxEx strMsg, vbExclamation, gstrSysName
                Exit Function
            End If
                
        End If
        '重新注册成功时需重新检查
        If blnReDo Then
            If Not CheckCertificate(strDBUser) Then Exit Function
        End If
264     CheckCertificate = True
        Exit Function
errH:
266     MsgBoxEx "CheckCertificate 错误出现在：" & CStr(Erl()) & "行," & Err.Description, vbExclamation, gstrSysName
End Function

Public Function Signature(ByVal strSource As String, ByVal strDBUser As String, lngCertID As Long, Optional strTimeStamp As String, _
    Optional objSignPic As StdPicture, Optional strTimeStampCode As String, Optional blnCheck As Boolean) As String
    '功能：生成签名,需要插入USB-Key
    '参数：strSource=签名原文(不包含时间戳信息)
    '      strDBUser=数据库用户名,用于取证书序列号来初始化证书
    '      blnCheck=False 调用CheckCert;True=不调用CheckCert 用于医嘱循环调用,减少不必要的检查提升签名速度
    '返回：1.签名信息
    '      2.lngCertID:返回签名使用的证书记录ID
    '      3.strTimeStamp:返回从时间戳服务器获取的时间戳(yyyy-MM-dd HH:mm:ss)，该时间戳信息已一并签名；没有时间戳服务器时返回空串
    '      4.objSignPic:返回签名图章(如果支持)
    '      5.strTimeStampCode=返回时间戳编码信息
        Dim rsTmp As New ADODB.Recordset
        Dim strSign As String, strData As String
        Dim strSQL As String, strPass As String
        Dim strBMP As String, strGIF As String
        Dim lngRet As Long, dCurDate As Date, dDate As Date
        Dim intCount As Integer '签名要花点时间，这个变量用于循环
        Dim arrData() As Byte, lngR As Long
        Dim Savetime As Single
        Dim NetONEXResult As Object ' NetONEX.TSAResponseX  '用于山东：时间戳返回对象
        Dim strTime As String
        Dim blnReDo As Boolean     'True -自动注册证书后需重新获取证书ID进行签名操作
        Dim blnRet As Boolean
        Dim lng证书ID As Long
        
10      On Error GoTo errH
20      LogWrite "Signature", "TIMER:" & Timer & "  原文:【" & strSource & "】"
30      lngCertID = 0
40      strTimeStamp = ""
50      Set objSignPic = Nothing
        
60      mlngCertId = 0
70      mstrTimeStamp = ""
80      mstrTimeStampCode = ""
90      Set mpicSign = Nothing

        '检查是否已注册
100     If Not blnCheck Then
110         strSQL = "Select B.ID,B.CertSN,B.CertDN,B.SignCert,B.EncCert,B.时间戳证书,C.ID as 人员ID,C.姓名,C.签名,C.身份证号 From 上机人员表 A,人员证书记录 B,人员表 C " & _
                " Where A.人员ID=B.人员ID And B.人员ID=C.ID And A.用户名=[1] Order by B.注册时间 Desc,B.ID Desc"
120         Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", strDBUser)
130         If rsTmp.EOF Then
140             MsgBoxEx "你的数字证书还没有注册，不能完成电子签名。", vbInformation, gstrSysName
150             Exit Function
160         End If
            
            '--获取当前操作员信息(北京CA会用到)----------------------------------------
170         mUserInfo.strID = Trim(rsTmp!人员ID & "")
180         mUserInfo.strEncCert = Trim(rsTmp!EncCert & "")
190         mUserInfo.strName = Trim(rsTmp!姓名 & "")
200         mUserInfo.strSignName = Trim(rsTmp!签名 & "")
210         mUserInfo.strUserID = Trim(rsTmp!身份证号 & "")
220         mUserInfo.strCertSN = Trim("" & rsTmp!CertSN)
230         mUserInfo.strCert = Trim("" & rsTmp!signCert)
240         mUserInfo.strTSCert = Trim("" & rsTmp!时间戳证书)
250         mUserInfo.strCertDN = Trim("" & rsTmp!CertDn)
260         mUserInfo.lngCertID = rsTmp!id
270     End If
280     lng证书ID = mUserInfo.lngCertID
        '-------------------------------------------
        
290     If gintCA = CA_辽宁省 Then
300         With gobjLNCA
                '1.客户端证书初始化(可重复进行)
310             If Not mblnLNCAinit Or mstrCurrUser <> strDBUser Then
                    Dim lngReturn As Long, intRt As Integer
320                 lngReturn = -1
330                 intRt = 0
340                 If mUserInfo.strCertSN = "" Then
350                     MsgBoxEx "当前用户" & strDBUser & "的SN为空，请重新注册！"
360                     mstrCurrUser = ""
370                     mblnLNCAinit = False
380                     Exit Function
390                 End If
400                 Do While intRt < 3
410                     lngReturn = .InitClientCertByCertSN(mUserInfo.strCertSN)
420                     intRt = intRt + 1
430                     If lngReturn = 0 Then Exit Do
440                     Savetime = Timer
450                     While Timer < Savetime + 2
460                         DoEvents
470                     Wend
480                 Loop
490                 If lngReturn <> 0 Then
500                     MsgBoxEx "数字证书初始化失败，不能完成电子签名！" & vbCrLf & vbCrLf & "错误编号:" & .errorCode & vbCrLf & "错误描述:" & .ErrorMessage, vbExclamation, gstrSysName
510                     mstrCurrUser = ""
520                     mblnLNCAinit = False
530                     Exit Function
540                 End If
550                 mstrCurrUser = strDBUser
560                 mblnLNCAinit = True
570             End If
                '2.生成签名(包含本地时间,不包含原文,证书),使用Ekey时会弹出窗口要求输入Ping码(密码)
580             strSign = .SignDataEx(strSource, "1.3.14.3.2.26", 0, 0, "", 1, "")
590             If strSign = "" Then
600                 MsgBoxEx "电子签名操作失败！" & vbCrLf & vbCrLf & "错误编号:" & .errorCode & vbCrLf & "错误描述:" & .ErrorMessage, vbExclamation, gstrSysName
610                 Exit Function
620             End If
630         End With
640     ElseIf gintCA = CA_广西省 Then
            '检查密钥盘及签名
650         Set mCerts = miTrusPTA.MyCertificates
660         If mCerts.Count > 0 Then
670             Set mCert = mCerts(1) '以第1个为准
680             If mCert.SerialNumber <> mUserInfo.strCertSN Then
690                 MsgBoxEx "当前数字证书与注册的不相符，不能完成电子签名。", vbExclamation, gstrSysName
700                 Exit Function
710             End If

                '检查证书是否过期
720             dCurDate = gobjComLib.zlDatabase.Currentdate
730             dDate = mCert.ValidTo
740             If Int(dDate) <= Int(dCurDate) Then
750                 MsgBoxEx "用户数字证书已于 " & Format(dDate, "yyyy-MM-dd") & " 失效，不能完成电子签名！", vbExclamation, gstrSysName
760                 Exit Function
770             Else
780                 dCurDate = DateAdd("d", 30, dCurDate)
790                 If Int(dDate) <= Int(dCurDate) And Not gblnShow Then
800                     MsgBoxEx "数字证书即将于" & Format(dDate, "yyyy-MM-dd") & " 失效，请注意及时更新证书！", vbInformation, gstrSysName
810                     gblnShow = True
820                 End If
830             End If

                '获取时间戳
840             If Not mobjTimeStamp Is Nothing Then
                    '返回时间串：yyyy-MM-dd HH:mm:ss，否则返回"0"
850                 strSQL = mobjTimeStamp.CATime()
860                 If IsDate(strSQL) Then
870                     strTimeStamp = Format(strSQL, "yyyy-MM-dd HH:mm:ss")
880                     strSource = strSource & strTimeStamp
890                 End If
900             End If

                '进行签名
910             strSign = mCert.SignMessage(strSource, OUTPUT_BASE64) '返回长度为2568的字符
            
920         Else
930             MsgBoxEx "没有发现有效的数字证书，不能完成电子签名。", vbExclamation, gstrSysName
940             Exit Function
950         End If
960     ElseIf gintCA = CA_北京CA广西 Then
970         blnRet = BJCAGX_Sign(Trim(mUserInfo.strCertSN), strSource, strSign, strTimeStamp, strTimeStampCode, blnReDo)
980         If Not blnRet And Not blnReDo Then Exit Function
990     ElseIf gintCA = CA_重庆市 Then
1000        If mobjCQCA Is Nothing Then
1010            MsgBoxEx "重庆CA中心的电子签名客户端环境未正确安装。", vbExclamation, gstrSysName
1020            Exit Function
1030        End If
            
1040        If mblnUSETSA Then
1050            If Not TimesWEB_Tamp(strSource, strTimeStamp) Then Exit Function
1060        End If
            '以签名源文的MD5值作为业务唯一标识
1070        strSign = mobjCQCA.WF_SignData(Md5_String_Calc(strSource & strTimeStamp), strSource & strTimeStamp, mUserInfo.strCertSN)

1080        If Len(strSign) < 10 Then
                'MsgboxEx "签名失败，错误代码：" & strSign, vbCritical, gstrSysName
1090            strSign = ""
1100        End If
1110    ElseIf gintCA = CA_山东省 Then
1120        If mobjSDCA Is Nothing Then
1130            MsgBoxEx "山东CA中心的电子签名控件未正确安装。", vbExclamation, gstrSysName
1140            Exit Function
1150        End If

1160        If CheckCurrUserPass(strDBUser, strPass) = False Then Exit Function

1170        lngRet = mobjSDCA.ReadCert(mstrSDcer, 2, strPass, mstrSDcer, 2, strPass)
1180        If lngRet <> 0 Then
1190            MsgBoxEx "错误" & lngRet & "：" & GetSDError(lngRet), vbExclamation, gstrSysName
1200            Exit Function
1210        End If

1220        If mobjSDCA.SignCertSN <> mUserInfo.strCertSN Then
1230            MsgBoxEx "当前数字证书与注册的不相符，不能完成电子签名。", vbExclamation, gstrSysName
1240            Exit Function
1250        End If

            '签名
1260        lngRet = mobjSDCA.SignData(mstrSDpfx, strPass, strSource, Len(strSource))
1270        If lngRet = 0 Then
1280            strSign = mobjSDCA.getconten() '获得密文的Base64码
1290        Else
1300            MsgBoxEx "错误" & lngRet & "：" & GetSDError(lngRet), vbExclamation, gstrSysName
1310            Exit Function
1320        End If
            '启用时间戳   新用户
1330        If UBound(Split(mstrPara, "|")) = 2 Then
1340            If Val(Split(mstrPara, "|")(2)) = 1 Then
1350                Set NetONEXResult = mNetONEX.TSACreate(strSource)
1360                If Not NetONEXResult Is Nothing Then
1370                    strTimeStampCode = NetONEXResult.ToBASE64() '时间戳Base64
1380                    strTime = NetONEXResult.TimeStamp
1390                    strTimeStamp = GetTimes(strTime)
1400                End If
1410            End If
1420        End If
            
            '获取签名图章
1430        If Not mobjSDImg Is Nothing Then
1440            strBMP = mobjFile.GetSpecialFolder(TemporaryFolder).Path & "\" & mobjFile.GetTempName
1450            strGIF = mobjFile.GetSpecialFolder(TemporaryFolder).Path & "\" & mobjFile.GetTempName

                'picindex=KEY中图章序列号，现KEY只支持存放两张图章，0-图章一，1-图章二
                'rgbflag=图章前景色：0-红色，1-蓝色，2-黑色
1460            lngRet = mobjSDCA.ShowSinglePic(strPass, 0, strBMP, 0) '文件证书方式调用有点慢
1470            If lngRet = 0 Then
1480                lngRet = mobjSDImg.Bmp2TransparentGif(strBMP, strGIF, 0)
1490                If lngRet = 0 Then
1500                    Set objSignPic = LoadPicture(strGIF)
1510                End If
1520            End If

1530            If mobjFile.FileExists(strBMP) Then mobjFile.DeleteFile strBMP, True
1540            If mobjFile.FileExists(strGIF) Then mobjFile.DeleteFile strGIF, True
1550        End If
1560    ElseIf gintCA = CA_吉林中心医院 Then

1570        If DateDiff("s", mLastCheckUSBTime, Now) > 3 Or strDBUser <> mstrCurrUser Then
                '超时，要验证是否插入USB
1580            If CheckCertificate(strDBUser) = False Then Exit Function '经电话联系陈东并确认，此处应传strDbUser
1590        End If

1600        arrData = PackBytes(strSource)
1610        strSign = mobjJLClient.DetachSign("", arrData(0), UBound(arrData) + 1)

1620        lngR = mobjJLClient.GetErrorCode
1630        If lngR <> 0 Then
1640            MsgBoxEx lngR & ":" & mobjJLClient.GetErrorMessage(lngR), vbExclamation, gstrSysName
1650            Exit Function
1660        End If

            '使CHECKUSB时间超时，好重新调用检查
1670        mLastCheckUSBTime = Now - 0.01
1680    ElseIf gintCA = CA_吉林省医院 Then
1690        If mobjJLSYY Is Nothing Then
1700            MsgBoxEx "电子签名控件未正确安装。", vbExclamation, gstrSysName
1710            Exit Function
1720        End If
1730        If DateDiff("s", mLastCheckUSBTime, Now) > 3 Or strDBUser <> mstrCurrUser Then
                '超时，要验证是否插入USB
1740            If CheckCertificate(strDBUser) = False Then Exit Function '经电话联系陈东并确认，此处应传strDbUser
1750        End If
1760        strSign = mobjJLSYY.SignString(strSource, True)
1770        If strSign = "" Then
1780            MsgBoxEx mobjJLSYY.GetLastError, vbExclamation, gstrSysName
1790            Exit Function
1800        End If
1810        mLastCheckUSBTime = Now - 0.01
1820    ElseIf gintCA = CA_准格尔医院 Then
1830        If Not ZGRYY_Sign(Trim$(mUserInfo.strCertSN), strSource, strSign, strTimeStamp) Then Exit Function
1840    ElseIf gintCA = CA_广东 Then
1850        If Not GDCA_Sign(Trim(mUserInfo.strCertSN), strSource, strSign, strTimeStamp) Then
1860            MsgBoxEx "签名失败！" & strSign, vbExclamation, gstrSysName
1870            Exit Function
1880        End If
1890    ElseIf gintCA = CA_吉林安信 Then
1900        blnRet = ANXIN_Sign(strSource, strSign, strTimeStamp, strTimeStampCode, blnReDo)
1910        If Not blnRet And Not blnReDo Then Exit Function
1920    ElseIf gintCA = CA_内蒙古 Then
1930        blnRet = NMG_Sign(strSource, strSign, strTimeStamp, strTimeStampCode, blnReDo, blnCheck)
1940        If Not blnRet And Not blnReDo Then Exit Function
1950    ElseIf gintCA = CA_北京 Then
1960        blnRet = BJCA_Sign(strSource, strSign, strTimeStamp, strTimeStampCode, blnReDo)
1970        If Not blnRet And Not blnReDo Then Exit Function
1980    ElseIf gintCA = CA_北京CA四川 Then
1990        If Not BJCASC_Sign(Trim(mUserInfo.strCertSN), strSource, strSign, strTimeStamp) Then Exit Function
2000    ElseIf gintCA = CA_北京CA湖北 Then
2010        If Not BJCAHB_Sign(Trim(mUserInfo.strCertSN), strSource, strSign, strTimeStamp, strTimeStampCode) Then Exit Function
2020    ElseIf gintCA = CA_北京CA辽宁 Then
2030        If Not BJCALN_Sign(Trim(mUserInfo.strCertSN), strSource, strSign, strTimeStamp) Then Exit Function
2040    ElseIf gintCA = CA_上海CA Then
2050        blnRet = SHCA_Sign(Trim(mUserInfo.strCertSN), strSource, strSign, strTimeStamp, blnReDo)
2060        If Not blnRet And Not blnReDo Then Exit Function
2070    ElseIf gintCA = CA_江苏CA Then
2080        If Not JSCA_Sign(Trim(mUserInfo.strCertSN), strSource, strSign, strTimeStamp, strTimeStampCode) Then Exit Function
2090    ElseIf gintCA = CA_新疆CA Then
2100        If Not XJCA_Sign(Trim(mUserInfo.strCertSN), strSource, strSign) Then Exit Function
2110    ElseIf gintCA = CA_北京CA江苏 Then
2120        blnRet = BJCAJS_Sign(Trim(mUserInfo.strCertSN), strSource, strSign, strTimeStamp, strTimeStampCode, blnReDo)
2130        If Not blnRet And Not blnReDo Then Exit Function
2140    ElseIf gintCA = CA_河北CA邯郸 Then
2150        blnRet = HBCA_Sign(strSource, strSign, strTimeStamp, strTimeStampCode, blnReDo)
2160        If Not blnRet And Not blnReDo Then Exit Function
2170    ElseIf gintCA = CA_河南CA商丘 Then
2180        blnRet = mobjESign.HNCA_Sign(Trim(mUserInfo.strCertSN), strSource, strSign, strTimeStamp, strTimeStampCode, blnReDo)
2190        If Not blnRet And Not blnReDo Then Exit Function
2200    ElseIf gintCA = CA_辽宁三院 Then
2210        blnRet = LNCA_Sign(strSource, strSign, strTimeStamp, strTimeStampCode, blnReDo, blnCheck)
2220        If Not blnRet And Not blnReDo Then Exit Function
2230    ElseIf gintCA = CA_湖北 Then
2240        If Not HUBEI_Sign(Trim(mUserInfo.strCertSN), strSource, strSign, strTimeStamp, strTimeStampCode) Then Exit Function
2250    ElseIf gintCA = CA_深圳 Then
2260        If Not SZCA_Sign(mUserInfo.strCertSN, strSource, strSign, strTimeStamp, strTimeStampCode) Then Exit Function
2270    ElseIf gintCA = CA_网证通 Then
2280        If Not WZT_Sign(strSource, strSign, strTimeStamp, strTimeStampCode) Then Exit Function
2290    ElseIf gintCA = CA_测试 Then
            '签名前检查（证书号，密码,效期等）
2300        If Not CheckCertificate(strDBUser) Then Exit Function
            '时间戳
2310        strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            '时间戳信息
2320        strTimeStampCode = "AAAA_TEST_CODE_INFO"
            '签名
2330        strSign = Md5_String_Calc(strSource & mUserInfo.strCertSN)
            '签名图片
2340        If mobjFile.FileExists(mTestSign.strSignImage) Then Set objSignPic = LoadPicture(mTestSign.strSignImage)
2350    End If
        
        '重新注册绑定后需重新获取证书ID
2360    If blnReDo Then
2370        strSign = Signature(strSource, strDBUser, lng证书ID, strTimeStamp, objSignPic, strTimeStampCode)
2380        If strSign = "" Then Exit Function
2390    End If
        
2400    lngCertID = lng证书ID
2410    Signature = strSign

2420    mlngCertId = lngCertID
2430    mstrTimeStamp = strTimeStamp
2440    mstrTimeStampCode = strTimeStampCode
2450    Set mpicSign = objSignPic
2460    LogWrite "Signature", "TIMER:" & Timer & "  签名值:【" & strSign & "】"
2470    LogWrite "Signature", "TIMER:" & Timer & "  时间戳信息:【" & strTimeStampCode & "】"
        
2480    Exit Function
errH:
2490    MsgBoxEx "Signature 错误出现在：" & CStr(Erl()) & "行,当前用户" & strDBUser & "," & Err.Description, vbExclamation, gstrSysName
End Function

Public Function VerifySignature(ByVal strSource As String, ByVal dblSignID As Double, ByVal intSignType As Integer, _
        Optional ByVal strSignValue As String, Optional ByVal strTStamp As String, Optional ByVal strTStampCode As String) As Boolean
'功能：验证签名,不需要插入USB-Key
'参数：strSource=签名原文(不包含时间戳信息),之所以传入是为了使原文生成规则在HIS程序中完成。
'      dblSignID=可以唯一确定一次签名信息的ID;RIS接口调用传【人员证书记录.ID】
'      intSignType=签名数据类型:0-RIS接口调用(此时lng签名ID传人的是人员证书记录的ID),1-医嘱,2-电子病历,3-药品收发记录,4-检验,
'                   5-老版护理,6-新版护理,7-血库,8-Pacs文档编辑器
'      intSignType=0 RIS接口调用,以下传人参数有效
'      strSignValue=签名时返回的签名值
'      strTStamp=签名时返回的时间串（格式："yyyy-MM-dd HH:mm:ss"）
'      strTStampCode=签名时返回的时间戳Base64信息
'返回：签名验证是否成功；成功时,已提示详细信息(为了可能显示一些必要信息)
    Dim rsTmp As New ADODB.Recordset
        Dim lng证书ID As Long, strSQL As String, str时间戳信息 As String
        Dim strData As String, lngRet As Long
        Dim strPass As String
        Dim bSuccess As Boolean
        Dim lngR As Long, strMsg As String, arrData() As Byte
        Dim str签章信息 As String
        Dim str签名信息 As String
        
        
        On Error GoTo errH
        LogWrite "VerifySignature", "原文:【" & strSource & "】"
        '读取原始签名相关信息
        If intSignType = 0 Then
            'RIS接口
            strSQL = "Select B.Certsn, B.Signcert, B.Enccert, B.时间戳证书,B.人员id From 人员证书记录 B Where b.Id = [1]"
            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", dblSignID)
100     ElseIf intSignType = 1 Then '医嘱签名数据
102         strSQL = "Select B.CertSN,B.SignCert,B.EncCert,B.时间戳证书,A.签名信息,A.时间戳,A.时间戳信息,B.人员ID From 医嘱签名记录 A,人员证书记录 B Where A.证书ID=B.ID And A.ID=[1]"
104         Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", dblSignID)
106     ElseIf intSignType = 2 Then '病历签名数据
108         strSQL = "Select 对象属性 From 电子病历内容 Where Id=[1] And 对象类型=8"
110         Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", dblSignID)
112         If rsTmp.EOF Then
114             MsgBoxEx "未发现有效的电子签名数据信息。", vbInformation, gstrSysName
                Exit Function
            End If
116         If Nvl(rsTmp!对象属性) <> "" Then
                lng证书ID = Val(Split(Nvl(rsTmp!对象属性), ";")(2)) '获取证书ID
                If UBound(Split(rsTmp!对象属性, ";")) >= 7 Then
                    str时间戳信息 = Split(rsTmp!对象属性, ";")(7)
                Else
                    str时间戳信息 = ""
                End If
            End If
118         rsTmp.Close

120         strSQL = "Select b.CertSN, b.SignCert, b.EncCert,B.时间戳证书,B.人员ID,a.要素值域 As 签名信息,a.要素单位 as 时间戳" & _
                " From (Select t.* From 电子病历内容 t Where t.Id = [1] And t.对象类型 = 8) a, 人员证书记录 b " & _
                " Where b.Id = [2]"
122         Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", dblSignID, lng证书ID)
124     ElseIf intSignType = 3 Then '药品签名数据
126         strSQL = "Select B.CertSN,B.SignCert,B.EncCert,B.时间戳证书,B.人员ID,A.签名信息,A.时间戳,A.时间戳信息 From 药品签名记录 A,人员证书记录 B Where A.证书ID=B.ID And A.ID=[1]"
128         Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", dblSignID)
130     ElseIf intSignType = 4 Then '检验签名数据
132         strSQL = "select * from (Select B.CertSN,B.SignCert,B.EncCert,A.签名信息,A.时间戳,A.时间戳信息 From 检验签名记录 A,人员证书记录 B Where A.证书ID=B.ID And A.检验标本ID=[1] order by a.签名时间  desc) where rownum = 1"
134         Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", dblSignID)
        ElseIf intSignType = 5 Then '老版护理
            strSQL = "Select B.CertSN,B.SignCert,B.EncCert,B.时间戳证书,B.人员ID,A.记录内容 AS 签名信息,A.项目名称 AS 时间戳,A.未记说明 As 时间戳信息 From 病人护理内容 A,人员证书记录 B Where A.项目ID=B.ID And A.ID=[1]"
135         Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", dblSignID)
        ElseIf intSignType = 6 Then '新版护理
            strSQL = "Select B.CertSN,B.SignCert,B.EncCert,B.时间戳证书,B.人员ID,A.记录内容 AS 签名信息,A.项目名称 AS 时间戳,A.未记说明 As 时间戳信息 From 病人护理明细 A,人员证书记录 B Where A.项目ID=B.ID And A.ID=[1]"
136         Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", dblSignID)
        ElseIf intSignType = 7 Then '血库
            strSQL = "Select B.CertSN,B.SignCert,B.EncCert,B.时间戳证书,B.人员ID,A.签名内容 AS 签名信息,A.签名时间 AS 时间戳,A.时间戳信息 From 血库签名 A,人员证书记录 B Where A.证书ID=B.ID And A.ID=[1]"
137         Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", dblSignID)
        ElseIf intSignType = 8 Then
            
            Dim lngStartIndex As Long
            Dim lngEndIndex As Long
            
            Dim strCertID As String             '证书ID
            lngStartIndex = InStr(strSource, "<signid>")
            lngEndIndex = InStr(strSource, "</signid>")
            strCertID = Mid(strSource, lngStartIndex + 8, lngEndIndex - lngStartIndex - 8)
            
            Dim strSignInfo As String           '签名信息
            lngStartIndex = InStr(strSource, "<signinfo>")
            lngEndIndex = InStr(strSource, "</signinfo>")
            strSignInfo = Mid(strSource, lngStartIndex + 10, lngEndIndex - lngStartIndex - 10)
            
            Dim strSignSource As String         '签名原文
            lngStartIndex = InStr(strSource, "<signsource>")
            lngEndIndex = InStr(strSource, "</signsource>")
            strSignSource = Mid(strSource, lngStartIndex + 12, lngEndIndex - lngStartIndex - 12)
            
            
            Dim strTimeStamp As String          '时间戳
            lngStartIndex = InStr(strSource, "<signtime>")
            lngEndIndex = InStr(strSource, "</signtime>")
            strTimeStamp = Mid(strSource, lngStartIndex + 10, lngEndIndex - lngStartIndex - 10)
            
            Dim strTimeStampCode As String      '时间戳编码
            lngStartIndex = InStr(strSource, "<signtimecode>")
            lngEndIndex = InStr(strSource, "</signtimecode>")
            strTimeStampCode = Mid(strSource, lngStartIndex + 14, lngEndIndex - lngStartIndex - 14)
            
            '签名信息可能超过四千，因此需要直接赋值
            str签名信息 = strSignInfo
            strSource = strSignSource
            
            strSQL = "select [1] as 签名信息, [2] as 签名原文, [3] as 时间戳, [4] as 时间戳信息, CertSn, SignCert, EncCert From 人员证书记录 Where ID=[5]"
            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", "", "", strTimeStamp, strTimeStampCode, strCertID)
        End If
        
138     If rsTmp.EOF Then
139         MsgBoxEx "未发现有效的电子签名数据信息。", vbInformation, gstrSysName
            Exit Function
        End If
        
        If intSignType > 0 Then
            If intSignType <> 8 Then
                '其他类型的签名验证
1400            strSignValue = "" & Nvl(rsTmp!签名信息)
            Else
                'pacs文档编辑器签名
                strSignValue = str签名信息
            End If
            
1401        strTStamp = "" & rsTmp!时间戳
            If intSignType = 2 Then
1402            strTStampCode = str时间戳信息
            Else
1403            strTStampCode = "" & rsTmp!时间戳信息
            End If
        End If
        
        If gintCA = CA_河北CA邯郸 Then
            str签章信息 = gobjComLib.Sys.Readlob(100, 14, rsTmp!人员ID & "," & rsTmp!CertSN, , 1)
        End If
        LogWrite "VerifySignature", "签名值:【" & strSignValue & "】"
        LogWrite "VerifySignature", "时间戳信息:【" & strTStampCode & "】"
        LogWrite "VerifySignature", "签章信息:【" & str签章信息 & "】"
140     If gintCA = CA_辽宁省 Then
142         With gobjLNCA
144             If intSignType = 2 Then     '电子病历的签名验证
                    '1.将签名时使用的证书作为每三方证书初始化
146                 If .InitThirdCert(rsTmp!signCert, rsTmp!EncCert) <> 0 Then
148                     MsgBoxEx "签名证书初始化失败，无法进行电子签名验证！" & vbCrLf & vbCrLf & "错误编号:" & .errorCode & vbCrLf & "错误描述:" & .ErrorMessage, vbExclamation, gstrSysName
                        Exit Function
                    End If

                    '2.验证签名数据:4-第三方的公钥证书
150                 strSource = .VerifySignEx(strSource, strSignValue, 4)
152                 If strSource = "" Then Exit Function '外部程序提示
                Else
                    '1.将签名时使用的证书作为每三方证书初始化
154                 If .InitThirdCert(rsTmp!signCert, rsTmp!EncCert) <> 0 Then
156                     MsgBoxEx "签名证书初始化失败，无法进行电子签名验证！" & vbCrLf & vbCrLf & "错误编号:" & .errorCode & vbCrLf & "错误描述:" & .ErrorMessage, vbExclamation, gstrSysName
                        Exit Function
                    End If

                    '2.验证签名数据:4-第三方的公钥证书
158                 strSource = .VerifySignEx(strSource, strSignValue, 4)
160                 If strSource <> "" Then
                        '4-第三方签名证书,5-第三方加密证书
162                     MsgBoxEx "验证成功，该电子签名数据有效：" & vbCrLf & vbCrLf & _
                            "签名证书公共名称:" & .ReadCertCN(4) & vbCrLf & _
                            "签名的客户端时间:" & .ReturnSigningTime(strSignValue), vbInformation, gstrSysName
                    Else
164                     MsgBoxEx "电子签名验证操作失败！" & vbCrLf & vbCrLf & "错误编号:" & .errorCode & vbCrLf & "错误描述:" & .ErrorMessage, vbExclamation, gstrSysName
                        Exit Function
                    End If
                End If
            End With
166     ElseIf gintCA = CA_广西省 Then
168         If intSignType = 1 Then '医嘱的签名验证
                '验证签名数据
                On Error Resume Next
170             Set mCert = miTrusPTA.VerifySignature(strSource & Format(strTStamp, "yyyy-MM-dd HH:mm:ss"), strSignValue, INPUT_BASE64)
172             If Err.Number <> 0 Or mCert Is Nothing Then
174                 Err.Clear
176                 MsgBoxEx "电子签名验证操作失败！", vbExclamation, gstrSysName
                    Exit Function
                End If
178             Err.Clear

180             MsgBoxEx "验证成功，该电子签名数据有效：" & vbCrLf & vbCrLf & "签名证书公共名称:" & mCert.CommonName & _
                    IIf(IsDate(strTStamp), vbCrLf & "时间戳:" & Format(strTStamp, "yyyy-MM-dd HH:mm:ss"), ""), vbInformation, gstrSysName
182         ElseIf intSignType = 2 Then '电子病历的签名验证
                '验证签名数据
                On Error Resume Next
184             Set mCert = miTrusPTA.VerifySignature(strSource & Format(strTStamp, "yyyy-MM-dd HH:mm:ss"), strSignValue, INPUT_BASE64)
186             If Err.Number <> 0 Or mCert Is Nothing Then
188                 Err.Clear: Exit Function
                End If
190             Err.Clear
            Else
                '验证签名数据
                On Error Resume Next
192             Set mCert = miTrusPTA.VerifySignature(strSource & Format(strTStamp, "yyyy-MM-dd HH:mm:ss"), strSignValue, INPUT_BASE64)
194             If Err.Number <> 0 Or mCert Is Nothing Then
196                 Err.Clear
198                 MsgBoxEx "电子签名验证操作失败！", vbExclamation, gstrSysName
                    Exit Function
                End If
200             Err.Clear
202             MsgBoxEx "验证成功，该电子签名数据有效：" & vbCrLf & vbCrLf & "签名证书公共名称:" & mCert.CommonName & _
                    IIf(IsDate(strTStamp), vbCrLf & "时间戳:" & Format(strTStamp, "yyyy-MM-dd HH:mm:ss"), ""), vbInformation, gstrSysName
            End If
204     ElseIf gintCA = CA_重庆市 Then
206         If mobjCQCA Is Nothing Then
208             MsgBoxEx "重庆CA中心的电子签名客户端环境未正确安装。", vbExclamation, gstrSysName
                Exit Function
            End If
            If mblnUSETSA Then
                '时间戳验证
                If Not verifyWEB_Timestamp(strSource) Then Exit Function
            End If
            '有证书信息传入，可以不插USBKey
210         If mobjCQCA.WF_VerifySigneData(rsTmp!signCert, strSource, strSignValue) Then
212             MsgBoxEx "验证成功，该电子签名数据有效。", vbInformation, gstrSysName
            Else
214             MsgBoxEx "电子签名验证操作失败！", vbExclamation, gstrSysName
                Exit Function
            End If
216     ElseIf gintCA = CA_山东省 Then
218         If mobjSDCA Is Nothing Then
220             MsgBoxEx "山东CA中心的电子签名控件未正确安装。", vbExclamation, gstrSysName
                Exit Function
            End If

222         If CheckCurrUserPass(mstrCurrUser, strPass) = False Then Exit Function

224         lngRet = mobjSDCA.ReadCert(mstrSDcer, 2, strPass, mstrSDcer, 2, strPass)
226         If lngRet <> 0 Then
228             MsgBoxEx "错误" & lngRet & "：" & GetSDError(lngRet), vbExclamation, gstrSysName
                Exit Function
            End If
            
            '时间戳验证
            If UBound(Split(mstrPara, "|")) = 2 Then
                If Val(Split(mstrPara, "|")(2)) = 1 Then
                    lngRet = mNetONEX.TSAVerify(strTStampCode)
                    If lngRet = 200 Then
                        'MsgboxEx ("验证成功！")
                    Else
                        MsgBoxEx "时间戳验证失败，该电子签名数据无效。", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
            '验证签名
            'rsTmp!SignCert
230         lngRet = mobjSDCA.VerifySign(mobjSDCA.getSignCert(), strSource, Len(strSource), strSignValue)
232         If lngRet <> 0 Then
234             MsgBoxEx "错误" & lngRet & "：" & GetSDError(lngRet), vbExclamation, gstrSysName
                Exit Function
            Else
236             MsgBoxEx "验证成功，该电子签名数据有效。", vbInformation, gstrSysName
            End If
238     ElseIf gintCA = CA_吉林中心医院 Then
240         If mobjJLServer Is Nothing Then
242             MsgBoxEx "服务器签名部件 JITClientCOMAPI.dll 没有安装。", vbExclamation, gstrSysName
            End If

244         arrData = PackBytes(strSource)
246         lngR = mobjJLServer.VerifyDetachedSignEx(strSignValue, arrData(0), UBound(arrData) + 1)
248         If lngR <> 0 Then
250             Call mobjJLServer.GetErrorMessageEx(lngR, strMsg)
252             MsgBoxEx lngR & ":" & strMsg, vbExclamation, "错误"
                Exit Function
            End If
254     ElseIf gintCA = CA_吉林省医院 Then
256         If mobjJLSYY Is Nothing Then
258             MsgBoxEx "电子签名控件未正确安装。", vbExclamation, gstrSysName
                Exit Function
            End If
260         bSuccess = mobjJLSYY.VerifyString(strSignValue, True, False, strSource)
262         If Not bSuccess Then
264             MsgBoxEx mobjJLSYY.GetLastError, vbExclamation, gstrSysName
                Exit Function
            End If
266     ElseIf gintCA = CA_准格尔医院 Then
 
268         If ZGRYY_VerifySign("" & rsTmp!CertSN, strSignValue, strSource, Format(strTStamp, "yyyy-MM-dd HH:mm:ss")) Then Exit Function
        ElseIf gintCA = CA_吉林安信 Then
            If ANXIN_VerifySign(strSource, strSignValue, strTStampCode) Then Exit Function
        ElseIf gintCA = CA_内蒙古 Then
            If NMG_VerifySign(strSource, strSignValue, strTStampCode) Then Exit Function
270     ElseIf gintCA = CA_广东 Then
272         If Not GDCA_VerifySign("" & rsTmp!CertSN, strSignValue, strSource) Then
274              MsgBoxEx "验证失败！", vbExclamation, gstrSysName
                Exit Function
            End If
276     ElseIf gintCA = CA_北京 Then
278         If Not BJCA_VerifySign("" & rsTmp!signCert, strSignValue, strSource, strTStampCode) Then Exit Function
280     ElseIf gintCA = CA_北京CA四川 Then
282         If Not BJCASC_VerifySign("" & rsTmp!CertSN, strSignValue, strSource) Then Exit Function
        ElseIf gintCA = CA_北京CA广西 Then
            If Not BJCAGX_VerifySign("" & rsTmp!CertSN, "" & rsTmp!signCert, strSignValue, strSource, strTStampCode) Then Exit Function
        ElseIf gintCA = CA_北京CA湖北 Then
            If Not BJCAHB_VerifySign("" & rsTmp!signCert, strSignValue, strSource, strTStampCode) Then Exit Function
        ElseIf gintCA = CA_北京CA辽宁 Then
            If Not BJCALN_VerifySign("" & rsTmp!CertSN, strSignValue, strSource) Then Exit Function
        ElseIf gintCA = CA_上海CA Then
            If Not SHCA_VerifySign("" & rsTmp!CertSN, strSignValue, strSource, "" & rsTmp!signCert) Then Exit Function
        ElseIf gintCA = CA_江苏CA Then
            If Not JSCA_VerifySign("" & rsTmp!signCert, strSignValue, strSource, Format(strTStamp, "yyyy-MM-dd HH:mm:ss"), strTStampCode) Then Exit Function
        ElseIf gintCA = CA_新疆CA Then
            If Not XJCA_VerifySign("" & rsTmp!signCert, strSignValue, strSource) Then Exit Function
        ElseIf gintCA = CA_北京CA江苏 Then
            If Not BJCAJS_VerifySign("" & rsTmp!CertSN, strSignValue, strSource, strTStampCode) Then Exit Function
        ElseIf gintCA = CA_河北CA邯郸 Then
            If Not HBCA_VerifySign(strSignValue, strSource, strTStampCode, "" & rsTmp!signCert, "" & rsTmp!时间戳证书, str签章信息) Then Exit Function
        ElseIf gintCA = CA_河南CA商丘 Then
            If Not mobjESign.HNCA_VerifySign("" & rsTmp!signCert, strSignValue, strSource, strTStampCode) Then Exit Function
        ElseIf gintCA = CA_辽宁三院 Then
            If Not LNCA_VerifySign("" & rsTmp!signCert, strSignValue, strSource) Then Exit Function
        ElseIf gintCA = CA_湖北 Then
            If Not HUBEI_VerifySign(strSignValue, strSource, strTStampCode) Then Exit Function
        ElseIf gintCA = CA_深圳 Then
            If Not SZCA_VerifySign("" & rsTmp!signCert, strSignValue, strSource, "" & rsTmp!时间戳证书, strTStampCode) Then Exit Function
        ElseIf gintCA = CA_网证通 Then
            If Not WZT_VerifySign(strSource, strSignValue) Then Exit Function
284     ElseIf gintCA = CA_测试 Then
              
286         If Md5_String_Calc(strSource & rsTmp!CertSN) = Trim(strSignValue) Then
288             MsgBoxEx "验证成功，该电子签名数据有效!"
            Else
290             MsgBoxEx "电子签名验证操作失败！"
                Exit Function
            End If
        End If
        
292     VerifySignature = True
        Exit Function
errH:
294     MsgBoxEx "Verifysignature 错误出现在：" & CStr(Erl()) & "行," & Err.Description, vbExclamation, gstrSysName

    '268     If gobjComLib.ErrCenter() = 1 Then Resume
    '270     Call gobjComLib.SaveErrLog
End Function

Private Sub Class_Initialize()
    mlngCertId = -1
    mstrTimeStamp = ""
    mstrTimeStampCode = ""
    Set mpicSign = Nothing
    glngSign = glngSign + 1
End Sub

Private Sub Class_Terminate()
    Dim lngR As Long, strMsg As String
    glngSign = glngSign - 1
    If glngSign > 0 Then Exit Sub
    
'    LogWrite "Class_Terminate", "签名部件释放,当前编号" & gintCA
    If gintCA = CA_辽宁省 Then
        Set gobjLNCA = Nothing
    ElseIf gintCA = CA_广西省 Then
        Set mobjTimeStamp = Nothing
    ElseIf gintCA = CA_重庆市 Then
        'Set mobjCQCA = Nothing
        Call TSAWEB_UnloadObj
    ElseIf gintCA = CA_山东省 Then
        Set mobjSDCA = Nothing
        Set mobjSDImg = Nothing
    ElseIf gintCA = CA_吉林中心医院 Then
        If Not mobjJLServer Is Nothing Then
            '释放服务器
            lngR = mobjJLServer.FinalizeServerConnectEx
            If lngR <> 0 Then
                Call mobjJLServer.GetErrorMessageEx(lngR, strMsg)
                MsgBoxEx lngR & ":" & strMsg, vbExclamation, "错误"
            End If
            Set mobjJLServer = Nothing
        End If
        If Not mobjJLClient Is Nothing Then
            Set mobjJLClient = Nothing
        End If
    ElseIf gintCA = CA_吉林省医院 Then
        If Not mobjJLSYY Is Nothing Then Set mobjJLSYY = Nothing
    ElseIf gintCA = CA_准格尔医院 Then
        Call ZGRYY_UnloadObj
    ElseIf gintCA = CA_广东 Then
        Call GDCA_UnloadObj
    ElseIf gintCA = CA_北京 Then
        Call BJCA_UloadObj
    ElseIf gintCA = CA_北京CA四川 Then
        Call BJCASC_UloadObj
    ElseIf gintCA = CA_北京CA广西 Then
        Call BJCAGX_UloadObj
    ElseIf gintCA = CA_北京CA湖北 Then
        Call BJCAHB_UloadObj
    ElseIf gintCA = CA_上海CA Then
        Call SHCA_UloadObj
    ElseIf gintCA = CA_江苏CA Then
        Call JSCA_UnloadObj
    ElseIf gintCA = CA_新疆CA Then
        Call XJCA_UnLoadObj
    ElseIf gintCA = CA_河北CA邯郸 Then
        Call HBCA_UnloadObj
    ElseIf gintCA = CA_河南CA商丘 Then
        Set mobjESign = Nothing
    ElseIf gintCA = CA_辽宁三院 Then
        Call LNCA_UnLoadObj
    ElseIf gintCA = CA_湖北 Then
        Call HUBEI_Unload
    ElseIf gintCA = CA_深圳 Then
        Call SZCA_UnloadObj
    ElseIf gintCA = CA_吉林安信 Then
        Call ANXIN_UnLoadObj
    ElseIf gintCA = CA_网证通 Then
        Call WZT_UnLoadObj
    End If
    
End Sub

Private Function CheckCurrUserPass(ByVal strDBUser As String, ByRef strPass As String) As Boolean
    '检查当前用户是否已经输入过密码，没有则调用窗体让用户输入后返回密码，已输入，则直接返回密码
    If mstrCurrUser = strDBUser And mstrCurrPass <> "" And mstrCurrUser <> "" Then
        '当前登录用户已输入过密码
        strPass = mstrCurrPass
        CheckCurrUserPass = True
    ElseIf mstrCurrUser = strDBUser And mstrCurrUser <> "" Then
        '还未输入过密码
        If Not frmPassword.ShowMe(strPass) Then Exit Function
        mstrCurrPass = strPass
        CheckCurrUserPass = True
    Else
        '不是当前登录用户，需要输密码
        If Not frmPassword.ShowMe(strPass) Then Exit Function
        mstrCurrPass = strPass
        CheckCurrUserPass = True
    End If
End Function

Public Function CertificateStoped(ByVal strUserName As String) As Boolean
    '判断指定人员的有效证书是否已停用，判断条件：有证书记录，且最近注册的被停用。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    strSQL = "select a.是否停用,a.注册时间 from 人员证书记录 a,人员表 b where a.人员id=b.id and b.姓名=[1] order by a.注册时间 desc  "
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "clsESign", strUserName)
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            If rsTmp!是否停用 = 1 Then
                CertificateStoped = True
            Else
                CertificateStoped = False
            End If
        Else
           CertificateStoped = False
        End If
    End If
End Function

Public Sub Setup(objFrm As Object, cnMain As ADODB.Connection, ByVal lngSys As Long)
'设置服务器参数
    Dim frmPara As New frmSetPara
    
        On Error GoTo errH
1000    Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
1100    Set gcnOracle = cnMain
        glngSys = lngSys
        gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
        
        Select Case gintCA
            Case CA_辽宁三院
                frmLNCASet.Show 1, objFrm
            Case CA_北京CA广西
                frmBJCAGX.Show 1, objFrm
            Case Else
1800            frmPara.Show 1, objFrm
        End Select
    Exit Sub
errH:
    MsgBoxEx "参数设置失败！错误出现在:第" & CStr(Erl()) & "行," & Err.Description, vbExclamation, "中联信息"
End Sub

Public Function SetEnabled(ByVal intCA As Integer) As Boolean
'功能:检查是否允许提供电子签名参数界面
    Dim blnDo As Boolean
    gintCA = intCA
    Select Case gintCA
    Case CA_湖北, CA_河北CA邯郸, CA_辽宁三院, CA_江苏CA, CA_北京, CA_新疆CA, CA_河南CA商丘, CA_深圳, CA_北京CA广西, CA_吉林安信
        blnDo = True
    Case CA_内蒙古, CA_网证通, CA_上海CA
        blnDo = True
    End Select
    
    SetEnabled = blnDo
End Function

Public Function GetESignType() As String
'功能:返回可选的接口字符串
    
    GetESignType = "0-不使用电子签名,1-辽宁省数字证书认证中心,2-广西省数字证书认证中心,3-重庆市数字证书认证中心,4-山东省数字证书认证中心," & _
     "5-吉大正元数字证书认证中心,6-国投安信数字证书认证中心,7-国投安信证书认证中心(内蒙),10-北京数字证书认证中心(河南)," & _
     "11-北京数字证书认证中心(四川),12-北京数字证书认证中心(广西),13-北京数字证书认证中心(湖北),14-北京数字证书认证中心(辽宁)," & _
     "15-上海数字证书认证中心(上海),16-江苏数字证书认证中心(江苏),17-新疆数字证书认证中心(新疆),18-北京数字证书认证中心(江苏)," & _
     "19-河北数字证书认证中心(邯郸),20-河南数字证书认证中心(商丘),21-辽宁数字证书认证中心(金秋),22-湖北数字证书认证中心,23-深圳数字证书认证中心," & _
     "24-吉林省安信数字证书认证中心,25-内蒙古数字证书认证中心,26-广东省数字证书认证中心(网证通)"
End Function

Public Function EnabledPatiSign() As Boolean
'是否支持患者签名，用于使用场合判断是否显示“患者签名”按扭
Dim blnDo As Boolean
    Select Case gintCA
        Case CA_辽宁三院
            If gudtPara.strSIGNIP = "" Then
                Call MsgBoxEx("没有配置手签服务器地址，请到【公共参数设置】先配置。", vbInformation, gstrSysName)
                blnDo = False
            Else
                blnDo = True
            End If
        Case CA_北京CA广西
                blnDo = True
        Case Else
            blnDo = False
    End Select
    EnabledPatiSign = blnDo
End Function

Public Function PenSignature(ByVal strSource As String, strName As String, strIdentifyNo As String, strOtherParms As String, _
    strSignInfo As String, strPenSignBase64 As String, objPenSignPic As StdPicture) As Boolean
'功能：获取手写签名图片、签名认证返回信息
'参数：strSource 签名源文
'        strName 病人姓名，缺省签名人
'        strIdentifyNo 病人身份证号，缺省签名人证件号
'        strOtherParms  为日后不变更参数数量，保留其它可能的参数
'        strSignInfo 签名返回认证信息
'        strPenSignBase64 返回手写签名图片BASE64编码
'        objPenSignPic 返回手写签名图片
Dim blnReturn As Boolean
    Select Case gintCA
        Case CA_辽宁三院
            If gudtPara.strSIGNIP = "" Then
                Call MsgBoxEx("没有配置手签服务器地址，请到【公共参数设置】先配置。", vbInformation, gstrSysName)
                Exit Function
            End If
            If gobjLNCAPenSign Is Nothing Then
                Set gobjLNCAPenSign = CreateObject("zlLNCAPenSign.clsPenSign")
            End If
            Call gobjLNCAPenSign.PenSignature(strSource, gudtPara.strSIGNIP, strName, strIdentifyNo, strOtherParms, strSignInfo, strPenSignBase64, objPenSignPic)
'            '---------测试用代码
'            strSignInfo = "aaaaa"
'            strPenSignBase64 = "bbbbb"
'            Set objPenSignPic = LoadPicture("c:\项目3.bmp")
'            '---------测试用代码
            blnReturn = True
        Case CA_北京CA广西
            If gobjGXCAPenSign Is Nothing Then
                Set gobjGXCAPenSign = CreateObject("zlGXCAPenSign.clsPenSign")
            End If
            blnReturn = gobjGXCAPenSign.PenSignature(strSource, Split(gudtPara.strSignURL, "|")(0), strName, strIdentifyNo, strOtherParms, strSignInfo, strPenSignBase64, objPenSignPic)
        Case Else
            blnReturn = False
    End Select
    
    PenSignature = blnReturn
    Exit Function
errH:
    MsgBoxEx "患者签名出错：" & Err.Description, vbExclamation, gstrSysName
End Function
Public Function ValidatePenSignature(ByVal strSource As String, strName As String, strIdentifyNo As String, strOtherParms As String, _
    strSignInfo As String, strPenSignBase64 As String) As Boolean
'功能：验证手写签名认证
'参数：strSource 签名源文
'        strName 签名人
'        strIdentifyNo 签名人证件号
'        strOtherParms  为日后不变更参数数量，保留其它可能的参数
'        strSignInfo 签名返回认证信息
'        strPenSignBase64 手写签名图片BASE64编码
Dim blnReturn As Boolean, strUrl As String
    Select Case gintCA
        Case CA_北京CA广西
            If gobjGXCAPenSign Is Nothing Then
                Set gobjGXCAPenSign = CreateObject("zlGXCAPenSign.clsPenSign")
            End If
            If UBound(Split(gudtPara.strSignURL, "|")) > 0 Then
                strUrl = Split(gudtPara.strSignURL, "|")(1)
            End If
            blnReturn = gobjGXCAPenSign.ValidatePenSignature(strSource, strUrl, strName, strIdentifyNo, strOtherParms, strSignInfo, strPenSignBase64)
        Case Else
            blnReturn = False
    End Select
    ValidatePenSignature = blnReturn
    Exit Function
End Function

Public Function EnabledVerifyPatiSign() As Boolean
'是否支持患者签名验证，用于使用场合判断是否显示“患者签名验证”按扭
Dim blnDo As Boolean
    Select Case gintCA
        Case CA_北京CA广西
            blnDo = True
        Case Else
            blnDo = False
    End Select
    EnabledVerifyPatiSign = blnDo
End Function

