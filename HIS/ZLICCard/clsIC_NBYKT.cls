VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIC_NBYKT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsCardDev

Private Const strFile As String = "C:\APPSOFT\NBYKT.INI"
Private gstr网址 As String, gstr端口号 As String, gstr目录 As String, gstr用户名 As String, gstr密码 As String, gstr完整地址 As String, gbln消息转发 As Boolean
Private gstr医院编码 As String, gstrIP As String, gstrMAC As String
Private mdomOutput As New MSXML2.DOMDocument

'------------------------------------------------------------
'以下是取本地MAC地址相关声明
Private Const MAX_ADAPTER_NAME_LENGTH As Long = 256
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH As Long = 128
Private Const MAX_ADAPTER_ADDRESS_LENGTH As Long = 8
Private Const ERROR_SUCCESS As Long = 0

Private Type MAC_ADDRESS_STRING
IpAddr(0 To 15) As Byte
End Type

Private Type MAC_MASK_STRING
IpMask(0 To 15) As Byte
End Type

Private Type MAC_ADDR_STRING
dwNext As Long
IpAddress As MAC_ADDRESS_STRING
IpMask As MAC_MASK_STRING
dwContext As Long
End Type

Private Type MAC_ADAPTER_INFO
dwNext As Long
ComboIndex As Long '保留
sAdapterName(0 To (MAX_ADAPTER_NAME_LENGTH + 3)) As Byte
sDescription(0 To (MAX_ADAPTER_DESCRIPTION_LENGTH + 3)) As Byte
dwAddressLength As Long
sMACAddress(0 To (MAX_ADAPTER_ADDRESS_LENGTH - 1)) As Byte
dwIndex As Long
uType As Long
uDhcpEnabled As Long
CurrentIpAddress As Long
IpAddressList As MAC_ADDR_STRING
GatewayList As MAC_ADDR_STRING
DhcpServer As MAC_ADDR_STRING
bHaveWins As Long
PrimaryWinsServer As MAC_ADDR_STRING
SecondaryWinsServer As MAC_ADDR_STRING
LeaseObtained As Long
LeaseExpires As Long
End Type

Private Declare Function GetAdaptersInfo Lib "iphlpapi.dll" (pTcpTable As Any, pdwSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dst As Any, src As Any, ByVal bcount As Long)
'------------------------------------------------------------


'------------------------------------------------------------
'以下是取密码相关声明
Private Declare Function GetPWD Lib "Authen.dll" (ByVal strUser As String) As String
'------------------------------------------------------------

Private Function LocalMACAddress() As String
    Dim cbRequired As Long
    Dim buff() As Byte
    Dim Adapter As MAC_ADAPTER_INFO
    Dim AdapterStr As MAC_ADDR_STRING
    Dim ptr1 As Long
    Dim sIPAddr As String
    Dim found As Boolean
    Dim iFound As Integer
    Dim sReturn As String
    
    iFound = 0
    sIPAddr = ""
    sReturn = ""
    Call GetAdaptersInfo(ByVal 0&, cbRequired)
    If cbRequired = 0 Then Exit Function
    ReDim buff(0 To cbRequired - 1) As Byte
    If GetAdaptersInfo(buff(0), cbRequired) <> ERROR_SUCCESS Then Exit Function
    
    '获取存放在buff()中的数据的指针
    ptr1 = VarPtr(buff(0))
    Do While (ptr1 <> 0)
        '将第一个网卡的数据转换到MAC_ADAPTER_INFO结构中
        CopyMemory Adapter, ByVal ptr1, LenB(Adapter)
        With Adapter
            'IpAddress.IpAddr成员给出了DHCP的IP地址
            Dim k As Long
            For k = 1 To .dwAddressLength
                sReturn = sReturn & Right("0" & Hex(AscB(MidB(.sMACAddress, k, 1))), 2) & "-"
            Next k
            sReturn = Left(sReturn, Len(sReturn) - 1) & vbCrLf
            ptr1 = .dwNext
        End With 'With Adapter
        '不再有网卡时，ptr1的值为0
    Loop 'Do While (ptr1 <> 0)
    '返回结果字符串
    LocalMACAddress = sReturn
End Function

Private Function zl_Ip_Address_FromOrc(Optional strDefaultIp_Address As String = "") As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:通过oracle获取的计算机的IP地址
    '入参:strDefaultIp_Address-缺省IP地址
    '出参:
    '返回:返回IP地址
    '编制:刘兴洪
    '日期:2009-01-21 11:08:47
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strIp_Address As String, strSQL As String
    Err = 0: On Error GoTo errHand:
     strSQL = "Select Sys_Context('USERENV', 'IP_ADDRESS') as Ip_Address From Dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "获取IP地址")
    If rsTemp.EOF = False Then
        strIp_Address = gobjCommFun.Nvl(rsTemp!Ip_Address)
    End If
    If strIp_Address = "" Then strIp_Address = strDefaultIp_Address
    If Replace(strIp_Address, " ", "") = "0.0.0.0" Then strIp_Address = ""
    zl_Ip_Address_FromOrc = strIp_Address
    Exit Function
errHand:
    
End Function

Private Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp >= lngLen Then
        '主要有空格引起的
        strTmp = ToVarchar(strCode, lngLen)
        lngTmp = LenB(StrConv(strTmp, vbFromUnicode))
    End If
    
    strTmp = strTmp & String(lngLen - lngTmp, strChar)
    '取掉最后半个字符
    'Rpad = Replace(strTmp, Chr(0), strChar)
    Rpad = strTmp
End Function

Private Function ToVarchar(ByVal varText As Variant, ByVal lngLength As Long) As String
'功能：将文本按Varchar2的长度计算方法进行截断
    Dim strText As String
    
    strText = IIf(IsNull(varText), "", varText)
    ToVarchar = StrConv(LeftB(StrConv(strText, vbFromUnicode), lngLength), vbUnicode)
    '去掉可能出现的半个字符
    ToVarchar = Replace(ToVarchar, Chr(0), "")
End Function

Private Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Dim varReturn As Variant
    varReturn = IIf(IsNull(varValue), DefaultValue, varValue)
    Nvl = Replace(varReturn, "'", "")
End Function

Private Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "")
'功能：打开记录集
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    
    rsTemp.Open strSQL, gcnConnect, adOpenStatic, adLockReadOnly
    Set rsTemp.ActiveConnection = Nothing
End Sub

Private Function clsCardDev_CloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    clsCardDev_CloseCard = True
End Function

Private Function clsCardDev_GetSpare() As Currency
    '
End Function

Private Function clsCardDev_initCard(Optional strInput As String, Optional strOut As String) As Boolean
    Dim strLine As String
    Dim strName As String
    Dim strValue As String
    Dim objStream As TextStream
    Dim objFileSys As New FileSystemObject
  
     Dim rsTemp  As New ADODB.Recordset
     Dim agcnconnect As New ADODB.Connection
     Dim conn As String, str用户 As String, str密码 As String
     
    On Error GoTo errHand
    
   
      '从配置文件中读取网址\目录及端口号
       str用户 = "system"
       str密码 = "abc123"
    gbln消息转发 = False
    If objFileSys.FileExists(strFile) Then
        Set objStream = objFileSys.OpenTextFile(strFile, ForReading)
        Do While Not objStream.AtEndOfStream
            strLine = objStream.ReadLine
            strName = Trim(Split(strLine, "=")(0))
            strValue = Trim(Split(strLine, "=")(1))
            Select Case strName
            Case "网址"
                gstr网址 = strValue
            Case "端口号"
                gstr端口号 = strValue
            Case "目录"
                gstr目录 = strValue
            Case "消息转发"
                gbln消息转发 = (Val(strValue) = 1)
            Case "服务器"
                  conn = Nvl(strValue, "orcl")
            Case "用户"
                  str用户 = Nvl(strValue, "system")
            Case "密码"
                  str密码 = Nvl(strValue, "abc123")
            End Select
        Loop
        objStream.Close
    Else
        gLastErr = "未进行宁波一卡通设置!"
        Exit Function
    End If
    
    
    
    
    If gcnConnect Is Nothing Then '如果没有传入gcnconnect则自己取，his错误
       
        If OraDataOpen(agcnconnect, conn, str用户, str密码, False) = False Then
                MsgBox "无法连接到中间库，请检查配置参数和中间库密码是否设置正确！"
                Exit Function
            End If
        Set gcnConnect = agcnconnect
    
    End If
    
    Dim strSQL As String
    rsTemp.CursorLocation = adUseClient
    If rsTemp.State = 1 Then rsTemp.Close
    strSQL = " Select 医院编码 From 一卡通目录 Where 名称=[1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "检查是否存在该病人的信息", "宁波一卡通")
    If rsTemp.RecordCount = 0 Then Exit Function
    If IsNull(rsTemp!医院编码) Then Exit Function
    gstr医院编码 = rsTemp!医院编码
    gstrIP = zl_Ip_Address_FromOrc()
    gstrMAC = Replace(Split(LocalMACAddress, vbCrLf)(0), "-", "")
      '取用户名
    gstr用户名 = gstr医院编码 & gstrMAC
    '取密码
    gstr密码 = GetPWD(gstr用户名)
'    gstr用户名 = "100100002170A67A35"
'    gstr密码 = "CCBA1FDD5E7C1CB3E342D50AB2EA28FD"
    
  
    
    gstr完整地址 = "http://" & gstr网址 & IIf(gstr端口号 <> "", ":" & gstr端口号, "") & "/" & gstr目录 & "/PersonInfoService?wsdl"
    clsCardDev_initCard = True
    Exit Function
errHand:
    gLastErr = Err.Description
End Function

Private Function clsCardDev_PaymentSwap(ByVal dblMoney As Double, dblBalance As Double, intCardType As Integer, ByVal lngHOSCode As Long, strCardNO As String, strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean
    '
End Function

Private Function clsCardDev_ReadCard(Optional strInput As String, Optional blnCancel As Boolean = False) As String
    '读卡
    clsCardDev_ReadCard = frmIdentifyNBYKT.ReadCard(gstr完整地址, gstr用户名, gstr密码, gbln消息转发)
    If clsCardDev_ReadCard = "" Then blnCancel = True
End Function

Private Function clsCardDev_ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean
    '
End Function

'保持属性值的局部变量
Private Sub clsCardDev_SetCard(Optional strInput As String, Optional strOut As String)
    frmSet_NBYKT.Show 1
End Sub

Private Function Get信息从表(ByVal lng病人ID As Long) As String
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    strSQL = "Select 信息值 From 病人信息从表 Where 病人ID=[1] And 信息名=[2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "读取病人信息从表", lng病人ID, "省")
    If rsTemp.RecordCount <> 0 Then
        Get信息从表 = Nvl(rsTemp!信息值)
    Else
        Get信息从表 = ""
    End If
    
    strSQL = "Select 信息值 From 病人信息从表 Where 病人ID=[1] And 信息名=[2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "读取病人信息从表", lng病人ID, "区")
    If rsTemp.RecordCount <> 0 Then
        Get信息从表 = Get信息从表 & "|" & Nvl(rsTemp!信息值)
    Else
        Get信息从表 = Get信息从表 & "|"
    End If
    
    strSQL = "Select 信息值 From 病人信息从表 Where 病人ID=[1] And 信息名=[2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "读取病人信息从表", lng病人ID, "街道")
    If rsTemp.RecordCount <> 0 Then
        Get信息从表 = Get信息从表 & "|" & Nvl(rsTemp!信息值)
    Else
        Get信息从表 = Get信息从表 & "|"
    End If
    
    strSQL = "Select 信息值 From 病人信息从表 Where 病人ID=[1] And 信息名=[2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "读取病人信息从表", lng病人ID, "单位地址")
    If rsTemp.RecordCount <> 0 Then
        Get信息从表 = Get信息从表 & "|" & Nvl(rsTemp!信息值)
    Else
        Get信息从表 = Get信息从表 & "|"
    End If
    
    strSQL = "Select 信息值 From 病人信息从表 Where 病人ID=[1] And 信息名=[2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "读取病人信息从表", lng病人ID, "手机号")
    If rsTemp.RecordCount <> 0 Then
        Get信息从表 = Get信息从表 & "|" & Nvl(rsTemp!信息值)
    Else
        Get信息从表 = Get信息从表 & "|"
    End If
    
    strSQL = "Select 信息值 From 病人信息从表 Where 病人ID=[1] And 信息名=[2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "读取病人信息从表", lng病人ID, "EMAIL")
    If rsTemp.RecordCount <> 0 Then
        Get信息从表 = Get信息从表 & "|" & Nvl(rsTemp!信息值)
    Else
        Get信息从表 = Get信息从表 & "|"
    End If
End Function

Private Function clsCardDev_UploadSwap(ByVal lng病人ID As Long, strOut As String) As Boolean
    Dim lngID As Long
    Dim strSQL As String
    Dim str病人信息从表 As String
    Dim strPatient As String
    Dim rsCard As New ADODB.Recordset
    Dim rsPatient As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '调用模块:就诊卡发放,门诊医生完成接诊,挂号结算,入院登记,出院登记
    '完成将病人信息以及卡号的更新上传
   If clsCardDev_initCard() = False Then
      Exit Function
      End If
      
   
 '   MsgBox "准备上传！"
    str病人信息从表 = Get信息从表(lng病人ID)
    
    If rsPatient.State = 1 Then rsPatient.Close
    strSQL = " Select 病人ID,住院次数 AS 主页ID,IC卡号,就诊卡号,操作类型,姓名,性别,出生日期,身份证号,其他证件,婚姻状况,家庭地址,家庭地址邮编," & _
                   "        家庭电话,工作单位,单位邮编,职业,单位电话,联系人姓名,联系人电话,一卡通建档时间 " & _
                   " From 病人信息" & _
                   " Where 病人ID=[1]"
    Set rsPatient = gobjDatabase.OpenSQLRecord(strSQL, "读取病人信息", lng病人ID)
    If rsPatient.RecordCount = 0 Then Exit Function
    '发卡记录最多只存在一条符合条件的记录
    If rsCard.State = 1 Then rsCard.Close
    strSQL = " Select A.旧卡号,A.旧卡类型,A.旧卡发卡医院,A.新卡号,A.发卡时间,A.旧卡发卡时间,A.旧卡明码 " & _
                " From 病人发卡记录 A" & _
                " Where Nvl(上传标志,0)=0 And 病人ID=[1] And 发卡时间=(Select MAX(发卡时间) From 病人发卡记录 Where 病人ID=[1])" & _
                " Order by 发卡时间"
    Set rsCard = gobjDatabase.OpenSQLRecord(strSQL, "读取病人发卡记录", lng病人ID)
    If rsCard.RecordCount = 0 Then Exit Function       '可能更新了病人信息的,所以不能退出
    
    '-----------------------------------------------
    '如果身份证或明码方式(旧卡类型=9),先删除旧卡数据
 '   MsgBox "准备上传旧卡数据"
    If Val(rsCard!旧卡类型) = 9 Then
    '    MsgBox "正在上传旧卡数据"
        '准备上传病人数据
        lngID = gobjDatabase.GetNextId("LOGID")
        strPatient = Rpad(gstr医院编码 & Format(gobjDatabase.CurrentDate, "yyyyMMdd") & Right("00000000" & CStr(lngID), 8), 30)       'LOGID
        strPatient = strPatient & Rpad(gstrIP, 15)                                                                  'IP
        strPatient = strPatient & Rpad(gstrMAC, 48)                                                                 'MAC
        strPatient = strPatient & gstr医院编码                                                                      '医院编码
        If IsNull(rsPatient!IC卡号) Or IsNull(rsPatient!操作类型) Then
            strPatient = strPatient & Rpad("", 22)             '档案号,新建档的病人传空
        Else
            strPatient = strPatient & Rpad(IIf(Val(rsPatient!IC卡号) = lng病人ID, "", rsPatient!IC卡号), 22)            '档案号,新建档的病人传空
        End If
        strPatient = strPatient & Rpad(rsPatient!姓名, 20)                                                          '姓名
        strPatient = strPatient & IIf(rsPatient!性别 = "男", 0, 1)                                                  '性别
        strPatient = strPatient & Format(rsPatient!出生日期, "yyyyMMdd")                                            '出生日期
        strPatient = strPatient & IIf(IsNull(rsPatient!身份证号), 9, 0)                                             '证件类型
        strPatient = strPatient & IIf(IsNull(rsPatient!身份证号), Rpad(Nvl(rsPatient!其他证件), 60), Rpad(Nvl(rsPatient!身份证号), 60))   '证件号码
        
        '取病人血型
        strSQL = "Select 信息值 From 病案主页从表 Where 病人ID=[1] And 主页ID=[2] And 信息名=[3]"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "取病人血型", rsPatient!病人ID, rsPatient!主页ID, "血型")
        If rsTemp.RecordCount <> 0 Then
            strPatient = strPatient & Left(rsTemp!信息值, 1)                                                        '血型
        Else
            strPatient = strPatient & " "                                                                           '血型
        End If
        
        strPatient = strPatient & IIf(rsPatient!婚姻状况 = "已婚", 0, 1)                                            '婚姻状况
        strPatient = strPatient & Rpad(Split(str病人信息从表, "|")(5), 60)                                           'EMAIL
        strPatient = strPatient & "2"                                                                   '卡类型
        strPatient = strPatient & Rpad(Trim(rsCard!旧卡号), 20)                                                     '卡号
        If Trim(rsCard!旧卡号) <> "" Then
            strPatient = strPatient & IIf(rsPatient!操作类型 = "明码", Rpad(Trim(rsCard!旧卡明码), 20), String(20, " "))                             '医保号/就诊卡明码
            strPatient = strPatient & rsCard!旧卡发卡时间                                                               '建卡时间
            strPatient = strPatient & Rpad(rsCard!旧卡发卡医院, 10)                                                              '旧卡发卡医院
        Else
            strPatient = strPatient & String(20, " ")                             '医保号/就诊卡明码
            strPatient = strPatient & String(14, " ")                                                              '建卡时间
            strPatient = strPatient & String(10, " ")                                                              '旧卡发卡医院
        End If
        strPatient = strPatient & IIf(Trim(rsCard!旧卡号) <> "", "4", "5")                                                                             '卡状态
        strPatient = strPatient & Format(gobjDatabase.CurrentDate, "yyyyMMddHHmmss")                                '操作时间
        strPatient = strPatient & Rpad(Nvl(rsPatient!家庭地址), 100)                                                '家庭地址
        strPatient = strPatient & Rpad(Split(str病人信息从表, "|")(0), 6)                                            '省/市
        strPatient = strPatient & Rpad(Split(str病人信息从表, "|")(1), 6)                                            '区
        strPatient = strPatient & Rpad(Split(str病人信息从表, "|")(2), 6)                                            '街道
        strPatient = strPatient & Rpad(Nvl(rsPatient!家庭地址邮编), 6)                                                  '家庭邮编
        strPatient = strPatient & Rpad(Nvl(rsPatient!家庭电话), 20)                                                 '家庭电话
        strPatient = strPatient & Rpad(Nvl(rsPatient!工作单位), 100)                                                '工作单位
        strPatient = strPatient & Rpad(Split(str病人信息从表, "|")(3), 100)                                          '单位地址
        strPatient = strPatient & Rpad(Nvl(rsPatient!单位邮编), 6)                                                  '单位邮编
        strPatient = strPatient & Rpad(Nvl(rsPatient!职业), 100)                                                    '职业
        strPatient = strPatient & Rpad(Nvl(rsPatient!单位电话), 20)                                                 '单位电话
        strPatient = strPatient & Rpad(Split(str病人信息从表, "|")(4), 20)                                           '手机号
        strPatient = strPatient & Rpad(Nvl(rsPatient!联系人姓名), 20)                                               '家属姓名
        strPatient = strPatient & Rpad(Nvl(rsPatient!联系人电话), 60)                                               '家属电话
        strPatient = strPatient & "1"                                                                               '操作类型
        'strPatient = strPatient & Format(gobjDatabase.CurrentDate, "yyyyMMddHHmmss")                               '操作时间
'        If Trim(rsCard!旧卡号) <> "" Then
'            strPatient = strPatient & Rpad(Trim(rsCard!旧卡号), 20)                                                           '旧卡号
'            strPatient = strPatient & rsCard!旧卡发卡医院                                                               '旧卡发卡医院
'        Else
            strPatient = strPatient & String(20, " ")                                                                   '旧卡号
            strPatient = strPatient & String(6, " ")                                                                    '旧卡发卡医院
'        End If
        strPatient = strPatient & String(20, " ")                                                                   '旧卡号
        strPatient = strPatient & String(6, " ")                                                                    '旧卡发卡医院
        strPatient = strPatient & String(20, " ")                                                                   '旧卡号
        strPatient = strPatient & String(6, " ")                                                                    '旧卡发卡医院
        strPatient = strPatient & String(20, " ")                                                                   '旧卡号
        strPatient = strPatient & String(6, " ")                                                                    '旧卡发卡医院
        strPatient = strPatient & String(20, " ")                                                                   '旧卡号
        strPatient = strPatient & String(6, " ")                                                                    '旧卡发卡医院
        strPatient = strPatient & rsPatient!一卡通建档时间                                                          '创建时间
        strPatient = strPatient & Format(gobjDatabase.CurrentDate, "yyyyMMddHHmmss")                                '更新时间
        Call 调用接口("putPersonInfo", strPatient)
    End If
    '-----------------------------------------------
    
    
    '-----------------------------------------------
    
    '准备上传病人数据
  '  MsgBox "准备上传新卡数据"
   ' If Not IsNull(rsCard!新卡号) Or (IsNull(rsCard!新卡号) And (rsCard!旧卡类型 = 2 Or rsPatient!操作类型 Like "异地*")) Then
     If Not IsNull(rsCard!新卡号) Or (IsNull(rsCard!新卡号) And rsCard!旧卡类型 = 2) Then
     
      ' MsgBox "正在上传新卡数据"
        lngID = gobjDatabase.GetNextId("LOGID")
        strPatient = Rpad(gstr医院编码 & Format(gobjDatabase.CurrentDate, "yyyyMMdd") & Right("00000000" & CStr(lngID), 8), 30)       'LOGID
        strPatient = strPatient & Rpad(gstrIP, 15)                                                                  'IP
        strPatient = strPatient & Rpad(gstrMAC, 48)                                                                 'MAC
        strPatient = strPatient & gstr医院编码                                                                      '医院编码
        If IsNull(rsPatient!IC卡号) Or (rsPatient!操作类型 Like "异地*") Or IsNull(rsPatient!操作类型) Then
            strPatient = strPatient & Rpad("", 22)             '档案号,新建档的病人传空
        Else
            strPatient = strPatient & Rpad(IIf(Val(rsPatient!IC卡号) = lng病人ID, "", rsPatient!IC卡号), 22)            '档案号,新建档的病人传空
        End If
        strPatient = strPatient & Rpad(rsPatient!姓名, 20)                                                          '姓名
        strPatient = strPatient & IIf(rsPatient!性别 = "男", 0, 1)                                                  '性别
        strPatient = strPatient & Format(rsPatient!出生日期, "yyyyMMdd")                                            '出生日期
        strPatient = strPatient & IIf(IsNull(rsPatient!身份证号), 9, 0)                                             '证件类型
        strPatient = strPatient & IIf(IsNull(rsPatient!身份证号), Rpad(Nvl(rsPatient!其他证件), 60), Rpad(Nvl(rsPatient!身份证号), 60))   '证件号码
        
        '取病人血型
        strSQL = "Select 信息值 From 病案主页从表 Where 病人ID=[1] And 主页ID=[2] And 信息名=[3]"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "取病人血型", rsPatient!病人ID, rsPatient!主页ID, "血型")
        If rsTemp.RecordCount <> 0 Then
            strPatient = strPatient & Left(rsTemp!信息值, 1)                                                        '血型
        Else
            strPatient = strPatient & " "                                                                           '血型
        End If
        
        strPatient = strPatient & IIf(rsPatient!婚姻状况 = "已婚", 0, 1)                                            '婚姻状况
        strPatient = strPatient & Rpad(Split(str病人信息从表, "|")(5), 60)                                           'EMAIL
        '新卡号为空说明病人本来就是通用就诊卡，不为空说明发了新的通用就诊卡，所以都从病人信息中取就诊卡号即可
        strPatient = strPatient & IIf(IsNull(rsCard!新卡号), rsCard!旧卡类型, IIf(rsCard!旧卡类型 = 0 Or rsCard!旧卡类型 = 1, rsCard!旧卡类型, 2))                                '卡类型
        strPatient = strPatient & Rpad(Nvl(rsPatient!就诊卡号, Nvl(rsCard!新卡号, rsCard!旧卡号)), 20)                                                    '卡号
        If InStr(1, "0,1", IIf(IsNull(rsCard!新卡号), rsCard!旧卡类型, 2)) <> 0 Then    '医保卡
            strPatient = strPatient & Rpad(Nvl(rsPatient!就诊卡号, Nvl(rsCard!新卡号, rsCard!旧卡号)), 20)
            strPatient = strPatient & Format(rsCard!发卡时间, "              ")                                         '建卡时间
            If rsPatient!操作类型 Like "异地*" Then
                strPatient = strPatient & Rpad("", 10)                                                            '发卡医院
            Else
                strPatient = strPatient & Rpad(gstr医院编码, 10)                                                            '发卡医院
            End If
        Else
            strPatient = strPatient & Rpad(Mid(rsPatient!就诊卡号, 1, 4) & Mid(rsPatient!就诊卡号, 11), 20)         '医保号/就诊卡明码
            If Not IsNull(rsCard!新卡号) Then
                strPatient = strPatient & Format(rsCard!发卡时间, "yyyyMMddHHmmss")                                         '建卡时间
                strPatient = strPatient & Rpad(gstr医院编码, 10)                                                            '发卡医院
            Else
                '跨院就诊
                strPatient = strPatient & rsCard!旧卡发卡时间                                         '建卡时间
                strPatient = strPatient & Rpad(rsCard!旧卡发卡医院, 10)                                                            '发卡医院
            End If
        End If
        strPatient = strPatient & "0"                                                                               '卡状态
        strPatient = strPatient & Format(gobjDatabase.CurrentDate, "yyyyMMddHHmmss")                                '操作时间
        strPatient = strPatient & Rpad(Nvl(rsPatient!家庭地址), 100)                                                '家庭地址
        strPatient = strPatient & Rpad(Split(str病人信息从表, "|")(0), 6)                                            '省/市
        strPatient = strPatient & Rpad(Split(str病人信息从表, "|")(1), 6)                                            '区
        strPatient = strPatient & Rpad(Split(str病人信息从表, "|")(2), 6)                                            '街道
        strPatient = strPatient & Rpad(Nvl(rsPatient!家庭地址邮编), 6)                                                  '家庭邮编
        strPatient = strPatient & Rpad(Nvl(rsPatient!家庭电话), 20)                                                 '家庭电话
        strPatient = strPatient & Rpad(Nvl(rsPatient!工作单位), 100)                                                '工作单位
        strPatient = strPatient & Rpad(Split(str病人信息从表, "|")(3), 100)                                          '单位地址
        strPatient = strPatient & Rpad(Nvl(rsPatient!单位邮编), 6)                                                  '单位邮编
        strPatient = strPatient & Rpad(Nvl(rsPatient!职业), 100)                                                    '职业
        strPatient = strPatient & Rpad(Nvl(rsPatient!单位电话), 20)                                                 '单位电话
        strPatient = strPatient & Rpad(Split(str病人信息从表, "|")(4), 20)                                           '手机号
        strPatient = strPatient & Rpad(Nvl(rsPatient!联系人姓名), 20)                                               '家属姓名
        strPatient = strPatient & Rpad(Nvl(rsPatient!联系人电话), 60)                                               '家属电话
        'strPatient = strPatient & Format(gobjDatabase.CurrentDate, "yyyyMMddHHmmss")                               '操作时间
'        If Trim(rsCard!旧卡号) <> "" And Not IsNull(rsCard!新卡号) Then
'            strPatient = strPatient & "1"                                                                           '操作类型
'            strPatient = strPatient & Rpad(Trim(rsCard!旧卡号), 20)                                                       '旧卡号
'            strPatient = strPatient & rsCard!旧卡发卡医院                                                           '旧卡发卡医院
'        Else
'            strPatient = strPatient & "0"                                                                           '操作类型
'            strPatient = strPatient & String(20, " ")                                                               '旧卡号
'            strPatient = strPatient & String(6, " ")                                                                '旧卡发卡医院
'        End If
        If rsCard!旧卡类型 = "9" Then
            strPatient = strPatient & "1"                                                                           '操作类型
            strPatient = strPatient & String(20, " ")                                                               '旧卡号
            strPatient = strPatient & String(6, " ")                                                                '旧卡发卡医院
        Else
            If Not IsNull(rsCard!新卡号) Or rsPatient!操作类型 Like "异地*" Then
                strPatient = strPatient & "0"                                                                           '操作类型
            Else
                strPatient = strPatient & "1"                                                                           '操作类型
            End If
            '如果是旧就诊卡,则需要填上旧卡号,其它情况不填写
            If rsCard!旧卡类型 = 2 And rsPatient!操作类型 = "就诊卡" Then
                strPatient = strPatient & Rpad(Trim(rsCard!旧卡号), 20)                                                       '旧卡号
                strPatient = strPatient & rsCard!旧卡发卡医院                                                           '旧卡发卡医院
            Else
                strPatient = strPatient & String(20, " ")                                                               '旧卡号
                strPatient = strPatient & String(6, " ")                                                                '旧卡发卡医院
            End If
        End If
        strPatient = strPatient & String(20, " ")                                                                   '旧卡号
        strPatient = strPatient & String(6, " ")                                                                    '旧卡发卡医院
        strPatient = strPatient & String(20, " ")                                                                   '旧卡号
        strPatient = strPatient & String(6, " ")                                                                    '旧卡发卡医院
        strPatient = strPatient & String(20, " ")                                                                   '旧卡号
        strPatient = strPatient & String(6, " ")                                                                    '旧卡发卡医院
        strPatient = strPatient & String(20, " ")                                                                   '旧卡号
        strPatient = strPatient & String(6, " ")                                                                    '旧卡发卡医院
        
        If rsPatient!操作类型 Like "异地*" Or _
            (rsPatient!操作类型 = "通用就诊卡" And Nvl(rsCard!旧卡发卡医院) = gstr医院编码) Or _
            rsPatient!操作类型 = "就诊卡" Then '新建档发卡，或旧卡换新卡
            strPatient = strPatient & Format(gobjDatabase.CurrentDate, "yyyyMMddHHmmss")                                '创建时间
        Else    '其它情况都传病人的建档时间
            strPatient = strPatient & Nvl(rsPatient!一卡通建档时间, Format(gobjDatabase.CurrentDate, "yyyyMMddHHmmss"))                                                          '创建时间
        End If
        
        strPatient = strPatient & Format(gobjDatabase.CurrentDate, "yyyyMMddHHmmss")                                '更新时间
        clsCardDev_UploadSwap = 调用接口("putPersonInfo", strPatient)
    End If
    '-----------------------------------------------
    
    '更新病人发卡记录的上传标志
    If clsCardDev_UploadSwap Then
        gcnConnect.Execute "zl_病人发卡记录_上传(" & lng病人ID & ")", , adCmdStoredProc
    End If
    
    clsCardDev_UploadSwap = True
    Exit Function
errHand:
    gLastErr = Err.Description
End Function

Private Function 调用接口(ByVal strFunction As String, ByVal strCreate As String) As Boolean
'    ----------------------------------------------------------------
    '功能描述   ：调用接口函数
    '编写人     ：朱玉宝
'    编写日期   ：2009-07-31
'    ----------------------------------------------------------------
    Dim strURL As String, strSoapRequest As String
    Dim str日期 As String, lng序列号 As Long, str错误信息 As String
    Dim objHttp As MSXML2.XMLHTTP
    On Error GoTo errHand
    
    Set objHttp = New MSXML2.XMLHTTP
    strURL = gstr完整地址 & "?op=" & strFunction
    
    strSoapRequest = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "utf-8" & Chr(34) & "?>" & _
                "<soapenv:Envelope xmlns:soapenv=" & Chr(34) & "http://schemas.xmlsoap.org/soap/envelope/" & Chr(34) & ">" & _
                "<soapenv:Header>" & _
                    "<ns:" & strFunction & " xmlns:ns=" & Chr(34) & "http://service.wondersgroup.com" & Chr(34) & ">" & _
                        "<ns:user>" & gstr用户名 & "</ns:user>" & _
                        "<ns:pwd>" & gstr密码 & "</ns:pwd>" & _
                    "</ns:" & strFunction & ">" & _
                "</soapenv:Header>" & _
                "<soapenv:Body>" & _
                    "<ns:" & strFunction & " xmlns:ns=" & Chr(34) & "http://service.wondersgroup.com" & Chr(34) & ">" & _
                        "<ns:FieldStr>" & strCreate & "</ns:FieldStr>" & _
                    "</ns:" & strFunction & ">" & _
                "</soapenv:Body>" & _
                "</soapenv:Envelope>"
       
    If gbln消息转发 = False Then
        objHttp.Open "post", strURL, False
        objHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        objHttp.setRequestHeader "Content-Length", Len(strSoapRequest)
        objHttp.setRequestHeader "SOAPAction", strURL
        
        '根据返回的状态信息来判断是否成功
        objHttp.send (strSoapRequest)
        If objHttp.status <> 200 Then
            MsgBox "返回信息：[" & objHttp.status & "]" & objHttp.responseText
            Exit Function
        End If
    Else
        '写入数据
        If Not SendRequest(str日期, lng序列号, strFunction, strURL, strSoapRequest) Then Exit Function
        
        '显示等待窗体
        If frmWait.SendRequest(str日期, lng序列号, str错误信息) = False Then
            If str错误信息 <> "" Then MsgBox str错误信息
            Exit Function
        End If
    End If
    
    '断点设置处
    Set mdomOutput = New MSXML2.DOMDocument
    If gbln消息转发 = False Then
        If mdomOutput.loadXML(objHttp.responseText) = False Then
            MsgBox "交易函数：" & strFunction & "，返回数据格式不正确！"
            Exit Function
        End If
    Else
        If mdomOutput.loadXML(str错误信息) = False Then
            MsgBox "交易函数：" & strFunction & "，返回数据格式不正确！"
            Exit Function
        End If
    End If
    
    调用接口 = True
    Exit Function
errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
End Function

Private Function SendRequest(str日期 As String, lng序列号 As Long, _
    ByVal strFuncName As String, ByVal strURL As String, ByVal strSoapRequest As String) As Boolean
    Dim blnTrans As Boolean
    Dim strRow As String
    Dim intRow As Integer, intCount As Integer
    On Error GoTo errHand
    '将待发送数据写入数据表
    
    str日期 = Format(gobjDatabase.CurrentDate, "yyyyMMdd")
    lng序列号 = gobjDatabase.GetNextId("消息转发")
    
    gcnConnect.BeginTrans
    blnTrans = True
    
    '插入主表
    gcnConnect.Execute "zl_消息主表_Insert('" & str日期 & "'," & lng序列号 & ",'" & strFuncName & "','" & strURL & "')", , adCmdStoredProc
    
    '插入待发送数据
    intCount = Len(strSoapRequest) \ 1000
    If Len(strSoapRequest) Mod 1000 <> 0 Then intCount = intCount + 1
    For intRow = 0 To intCount
        strRow = Mid(strSoapRequest, intRow * 1000 + 1, 1000)
        gcnConnect.Execute "zl_消息转发_Insert('" & str日期 & "'," & lng序列号 & "," & intRow + 1 & ",'" & strRow & "')", , adCmdStoredProc
    Next
    
    gcnConnect.CommitTrans
    blnTrans = False
    SendRequest = True
    Exit Function
errHand:
    If blnTrans Then gcnConnect.RollbackTrans
    MsgBox Err.Description
End Function

