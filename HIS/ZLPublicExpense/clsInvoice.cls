VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum gEM_BillType
    EM_收费收据 = 1
    EM_预交收据 = 2
    EM_结帐收据 = 3
    EM_挂号收据 = 4
    EM_医疗卡收据 = 5
    EM_退费收据 = 11
    EM_结帐红票 = 13
End Enum

Public Function zlGetInvoiceUserType(ByVal lngModule As Long, ByVal bytBillType As gEM_BillType, _
    ByVal lng病人ID As Long, ByVal lng主页Id As Long, _
    ByVal intInsure As Integer, ByRef strOutUseType As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的使用类别
    '入参:lngModule-模块号
    '     bytBillType-票种
    '     lng病人ID-病人ID
    '     lng主页ID-主页ID
    '     intInsure-险类
    '出参:strOutUseType-返回使用类别
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-27 10:14:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    
    On Error GoTo errHandle
    strOutUseType = ""
'    If Not zlStartFactUseType(bytBillType) Then zlGetInvoiceUserType = True: Exit Function
    
    strSql = "Select  Zl_Billclass([1],[2],[3]) as 使用类别 From Dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取票据使用类别", lng病人ID, lng主页Id, intInsure)
    strOutUseType = Nvl(rsTemp!使用类别)
    zlGetInvoiceUserType = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function GetDrawShareInvoiceRec(ByVal bytBillType As gEM_BillType) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定票种的共用票据批次
    '入参:bytBillType -票种
    '返回:rsOutDraw-返回共用票据集(ID,领用人,登记时间,开始号码,终止号码,剩余数量)
    '编制:刘尔旋
    '日期:2014-05-23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo errH
    If bytBillType = 1 Or bytBillType = 3 Then  '收费和结帐
        strSql = "" & _
        "   Select A.ID,nvl(M.编码,' ') as 使用类别编码,A.使用类别,A.领用人,A.登记时间,A.开始号码,A.终止号码,A.剩余数量 " & _
        "   From 票据领用记录 A,人员表 B,票据使用类别 M" & vbNewLine & _
        "   Where A.票种=[1] And A.使用方式=2 And A.剩余数量>0 And A.领用人=B.姓名" & _
        "           And A.使用类别=M.名称(+) " & _
        "           And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
        "   Order by 使用类别编码,剩余数量 Desc"
    ElseIf bytBillType = 5 Then
        '就诊卡
        strSql = "" & _
        "   Select A.ID,nvl(M.编码,' ') as 使用类别编码,M.ID as 使用类别ID,M.名称 as 使用类别,A.领用人,A.登记时间,A.开始号码,A.终止号码,A.剩余数量 " & _
        "   From 票据领用记录 A,人员表 B,医疗卡类别 M" & vbNewLine & _
        "   Where A.票种=[1] And A.使用方式=2 And A.剩余数量>0 And A.领用人=B.姓名" & _
        "           And to_number(nvl(A.使用类别,'0'))=M.ID(+) " & _
        "           And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
        "   Order by 使用类别编码,剩余数量 Desc"
    ElseIf bytBillType = 2 Then  '预交
        strSql = "" & _
        "   Select A.ID,to_number(nvl(A.使用类别,'0')) as 使用类别,A.领用人,A.登记时间,A.开始号码,A.终止号码,A.剩余数量 " & _
        "   From 票据领用记录 A,人员表 B" & vbNewLine & _
        "   Where A.票种=[1] And A.使用方式=2 And A.剩余数量>0 And A.领用人=B.姓名" & _
        "           And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
        "   Order by 使用类别,剩余数量 Desc"
    Else
        strSql = "" & _
        "   Select A.ID,A.使用类别,A.领用人,A.登记时间,A.开始号码,A.终止号码,A.剩余数量 " & _
        "   From 票据领用记录 A,人员表 B" & vbNewLine & _
        "   Where A.票种=[1] And A.使用方式=2 And A.剩余数量>0 And A.领用人=B.姓名" & _
        "           And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
        "   Order by 使用类别,剩余数量 Desc"
    End If
    Set GetDrawShareInvoiceRec = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, bytBillType)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlGetInvoiceShareID(ByVal lngModule As Long, _
    ByVal bytBillType As gEM_BillType, _
    Optional ByVal strUseType As String = "", _
    Optional ByRef lngShareInvoiceID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票的共享票据ID
    '入参:lngModule-模块号
    '     strUseType-使用类别
    '     bytBillType-票种
    '出参:lngShareInvoiceID-返回共用的发票ID
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-27 10:22:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, strShareTypeUseID As String
    Dim lng领用ID As Long, rsTemp As ADODB.Recordset, strSql As String
    Dim lng共用票据 As Long, lng使用票据 As Long, trFactType As String
    Dim strFactType As String
    
    lngShareInvoiceID = 0
    '因为Getpara就缓存了的,所以不用先用变量进行记录
    If lngModule = 1137 Then
        strFactType = "" & _
        Switch(bytBillType = EM_挂号收据, "共用挂号票据批次", _
                bytBillType = EM_结帐收据, "共用结帐票据批次", _
                bytBillType = EM_收费收据, "共用结帐票据批次", _
                bytBillType = EM_医疗卡收据, "共用医疗卡批次", _
                bytBillType = EM_预交收据, "共用预交票据批次", _
                True, "")
    Else
        strFactType = "" & _
        Switch(bytBillType = EM_挂号收据, "共用挂号票据批次", _
                bytBillType = EM_结帐收据, "共用结帐票据批次", _
                bytBillType = EM_收费收据, "共用收费票据批次", _
                bytBillType = EM_医疗卡收据, "共用医疗卡批次", _
                bytBillType = EM_预交收据, "共用预交票据批次", _
                True, "")
    End If
    If strFactType = "" Then Exit Function
    
    '格式:领用ID1,使用类别1|....
    strShareTypeUseID = Trim(gobjDatabase.GetPara(strFactType, glngSys, lngModule, "0"))
    
    varData = Split(strShareTypeUseID, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        lng领用ID = Val(varTemp(0))
        If Trim(varTemp(1)) = "" Then lng共用票据 = lng领用ID    '共用的.
        If Trim(varTemp(1)) = strUseType And lng领用ID <> 0 Then
            lng使用票据 = lng领用ID
        End If
    Next
    
    On Error GoTo errHandle
    '优先顺序
    '1.先使用
    '2.使用类别不区分的
    '3.具体使用类别的
    strSql = _
    "Select ID, 前缀文本, 开始号码, 终止号码, 剩余数量, 登记时间, 使用时间" & vbNewLine & _
    "From 票据领用记录" & vbNewLine & _
    "Where (ID =[1] or ID =[2]) And 剩余数量 > 0   " & vbNewLine & _
    "Order By Nvl(使用时间, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,使用类别 Desc, 开始号码"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "可用票据批次", lng共用票据, lng使用票据)
    
    If Not rsTemp.EOF Then lngShareInvoiceID = Val(Nvl(rsTemp!ID))
    Set rsTemp = Nothing
    zlGetInvoiceShareID = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlGetInvoicePrintFormat(ByVal lngModule As Long, _
    ByVal bytBillType As gEM_BillType, _
    Optional ByVal strUseType As String = "", _
    Optional ByRef intOutPrintFormat As Integer, _
    Optional ByVal bytInExseKind As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的打印格式
    '入参:lngModule-模块号
    '     strUseType-使用类别
    '     bytBillType-票种
    '出参:intOutPrintFormat-返回票据打印格式
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-27 10:22:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFactTypeFormat  As String, intPrintFormat As Integer, intPrintFormat1 As Integer
    Dim varData As Variant, varTemp As Variant, i As Integer
    On Error GoTo errHandle
    
    strFactTypeFormat = "" & _
    Switch(bytBillType = EM_挂号收据, "挂号发票格式", _
               bytBillType = EM_结帐收据, "住院结帐发票格式", _
               bytBillType = EM_收费收据, "收费发票格式", _
               bytBillType = EM_医疗卡收据, "医疗卡发票格式", _
               bytBillType = EM_预交收据, "预交发票格式", _
               bytBillType = EM_退费收据, "退费发票格式", _
               bytBillType = EM_结帐红票, "作废发票格式", _
               True, "")
               
    If bytBillType = EM_结帐收据 Or bytBillType = EM_收费收据 Then
        If bytInExseKind = 1 Then
            strFactTypeFormat = "门诊结帐发票格式"
        End If
        If bytInExseKind = 2 Then
            strFactTypeFormat = "住院结帐发票格式"
        End If
    End If
    
    intOutPrintFormat = 0
    '初始发票格式
    If bytBillType = EM_预交收据 Then
        '预交暂无格式
       intOutPrintFormat = 0: zlGetInvoicePrintFormat = True
       Exit Function
    End If
    
    strFactTypeFormat = Trim(gobjDatabase.GetPara(strFactTypeFormat, glngSys, lngModule, ""))
    '格式:使用类别1,格式1|使用类别2,格式2...
    varData = Split(strFactTypeFormat, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        intPrintFormat = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then intPrintFormat1 = intPrintFormat
        If Trim(varTemp(0)) = strUseType And intPrintFormat <> 0 Then
           intOutPrintFormat = intPrintFormat: Exit For
        End If
    Next
    If intOutPrintFormat = 0 And intPrintFormat1 <> 0 Then intOutPrintFormat = intPrintFormat1
    zlGetInvoicePrintFormat = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetInvoicePrintMode(ByVal lngModule As Long, _
    ByVal bytBillType As gEM_BillType, _
    Optional ByVal strUseType As String = "", _
    Optional ByRef intOutPrintMode As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的打印格式
    '入参:lngModule-模块号
    '     strUseType-使用类别
    '     bytBillType-票种
    '出参:intOutPrintMode-返回票据打印方式
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-27 10:22:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFacePrintMode  As String, intPrintMode As Integer, intPrintMode1 As Integer
    Dim varData As Variant, varTemp As Variant, i As Integer
    On Error GoTo errHandle
    If lngModule = 1137 Then
        strFacePrintMode = "" & _
        Switch(bytBillType = EM_挂号收据, "挂号发票打印方式", _
                   bytBillType = EM_结帐收据, "病人结帐打印", _
                   bytBillType = EM_收费收据, "病人结帐打印", _
                   bytBillType = EM_医疗卡收据, "医疗卡发票打印方式", _
                   bytBillType = EM_预交收据, "预交发票打印方式", _
                   bytBillType = EM_退费收据, "作废发票打印方式", _
                   bytBillType = EM_结帐红票, "作废发票打印方式", _
                   True, "")
    Else
        strFacePrintMode = "" & _
        Switch(bytBillType = EM_挂号收据, "挂号发票打印方式", _
                   bytBillType = EM_结帐收据, "病人结帐打印", _
                   bytBillType = EM_收费收据, "收费发票打印方式", _
                   bytBillType = EM_医疗卡收据, "医疗卡发票打印方式", _
                   bytBillType = EM_预交收据, "预交发票打印方式", _
                   bytBillType = EM_退费收据, "退费发票打印方式", _
                   bytBillType = EM_结帐红票, "作废发票打印方式", _
                   True, "")
    End If
    
    intOutPrintMode = 0
    
    strFacePrintMode = Trim(gobjDatabase.GetPara(strFacePrintMode, glngSys, lngModule, ""))
    
    intOutPrintMode = -1
    '格式:使用类别1,打印方式1|使用类别2,打印方式2...
    varData = Split(strFacePrintMode, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",,", ",")
        intPrintMode = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then intPrintMode1 = intPrintMode
        If Trim(varTemp(0)) = strUseType Then
            intOutPrintMode = intPrintMode: Exit For
        End If
    Next
    If intOutPrintMode < 0 Then intOutPrintMode = intPrintMode1
    zlGetInvoicePrintMode = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetInvoiceNoLength(ByVal bytBillType As gEM_BillType, Optional ByRef lngCardTypeID As Long) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据长度
    '返回:返回票据长度
    '编制:刘兴洪
    '日期:2014-09-30 10:52:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim intLen As Long, i As Integer
    Dim varTemp As Variant
    
    intLen = 7
    If bytBillType = EM_医疗卡收据 Then
        If lngCardTypeID <> 0 Then
            strSql = "select nvl(是否严格控制,0) as 卡号长度 from 医疗卡类别 where ID=[1] "
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "取医疗卡票号长度", lngCardTypeID)
            If Not rsTemp.EOF Then
                intLen = Val(Nvl(rsTemp!卡号长度))
            End If
        End If
        zlGetInvoiceNoLength = intLen
        Exit Function
    End If

    i = Switch(bytBillType = EM_挂号收据, 4, _
        bytBillType = EM_结帐收据, 3, _
        bytBillType = EM_收费收据, 1, _
        bytBillType = EM_医疗卡收据, 0, _
        bytBillType = EM_预交收据, 2, _
        True, 0)
    '位数分别:1-收费,2-预交,3-结帐,4-挂号
    varTemp = Split(gobjDatabase.GetPara(20, glngSys, , "7|7|7|7|7"), "|")
    If UBound(varTemp) >= i - 1 Then
        intLen = Val(varTemp(i - 1))
    End If
    zlGetInvoiceNoLength = intLen
End Function

Public Function zlGetInvoicePreperty(lngModule As Long, _
     ByVal bytBillType As gEM_BillType, _
     ByVal lng病人ID As Long, ByVal lng主页Id As Long, _
     ByVal intInsure As Integer, _
     ByRef objFactProperty As clsFactProperty, _
     Optional ByRef lngCardTypeID As Long, _
     Optional ByVal bln门诊预交 As Boolean, _
     Optional ByVal bytInExseKind As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定病人的发票信息,将相关信息返回给ObjFaceProperty对象
    '入参:lngModule-模块号
    '     lng病人ID-病人ID
    '     lng主页ID-主页ID
    '     intInsure-险类
    '     bytBillType-票种
    '     lngCardTypeID-卡类别ID(bytBillType=医疗卡时有效),主要获取该卡是否严格控制票据
    '     bln门诊预交-预交类型
    '     bytInExseKind-结帐类型,1-门诊;2-住院
    '出参:objFactProperty-返回发票对象
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-27 11:17:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, i As Integer, strSql As String, rsTemp As ADODB.Recordset
    Dim strUseType As String, lngShareID As Long, intPrintMode As Integer, intPrintFormat As Integer
    
    
    Set objFactProperty = New clsFactProperty
    
    objFactProperty.票种 = IIf(bytBillType = EM_结帐红票, EM_结帐收据, IIf(bytBillType = EM_退费收据, EM_收费收据, bytBillType))
    On Error GoTo errHandle
    '获取使用类别
    If bytBillType <> 2 Then
        Call zlGetInvoiceUserType(lngModule, objFactProperty.票种, lng病人ID, lng主页Id, intInsure, strUseType)
        objFactProperty.使用类别 = strUseType
        objFactProperty.启用使用类别 = zlStartFactUseType(objFactProperty.票种)
    Else
        objFactProperty.使用类别 = IIf(bln门诊预交, "1", "2")
        objFactProperty.启用使用类别 = True
    End If
    Call zlGetInvoiceShareID(lngModule, objFactProperty.票种, objFactProperty.使用类别, lngShareID)
    objFactProperty.共享批次ID = lngShareID
    
    Call zlGetInvoicePrintMode(lngModule, bytBillType, objFactProperty.使用类别, intPrintMode)
    objFactProperty.打印方式 = intPrintMode
    Call zlGetInvoicePrintFormat(lngModule, bytBillType, objFactProperty.使用类别, intPrintFormat, bytInExseKind)
    objFactProperty.打印格式 = intPrintFormat
    objFactProperty.票号长度 = 7
    If bytBillType = EM_医疗卡收据 Then
         objFactProperty.严格控制 = 0
        If lngCardTypeID <> 0 Then
            strSql = "select nvl(是否严格控制,0) as 严格控制,卡号长度 from 医疗卡类别 where ID=[1] "
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "取医疗卡是否严格控制", lngCardTypeID)
            If Not rsTemp.EOF Then
                objFactProperty.严格控制 = Val(Nvl(rsTemp!严格控制)) <> 0
                objFactProperty.票号长度 = Val(Nvl(rsTemp!卡号长度))
            End If
        End If
    Else
        i = Switch(bytBillType = EM_挂号收据, 4, _
               bytBillType = EM_结帐收据, 3, _
               bytBillType = EM_收费收据, 1, _
               bytBillType = EM_医疗卡收据, 0, _
               bytBillType = EM_预交收据, 2, _
               bytBillType = EM_退费收据, 1, _
               bytBillType = EM_结帐红票, 3, _
               True, 0)
        '位数分别:1-收费,2-预交,3-结帐,4-挂号
        objFactProperty.严格控制 = Val(Mid(gobjDatabase.GetPara(24, glngSys, , "00000"), i, 1)) = 1
        objFactProperty.票号长度 = zlGetInvoiceNoLength(objFactProperty.票种, lngCardTypeID)
    End If
    zlGetInvoicePreperty = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function zlCheckInvoiceIsUsed(lngModule As Long, _
       ByVal lng领用ID As Long, ByVal strInvoiceNO As String, ByRef blnOutUsed As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查发票是否已被使用
    '入参:lngModule -模块号
    '     lng领用ID-领用ID
    '     strInvoiceNo-检查的发票号
    '出参:blnOutUsed-true:表示使用;False-表示未使用
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-27 11:49:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    
    On Error GoTo errHandle
    blnOutUsed = False
    
    strSql = "Select 号码 From 票据使用明细 Where 领用ID=[1]  And 号码=[2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "检查发票是否已被使用", lng领用ID, strInvoiceNO)
    blnOutUsed = Not rsTemp.EOF
    
    zlCheckInvoiceIsUsed = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetInvoiceGroupID(ByVal strUserName As String, ByVal bytKind As Byte, ByVal intNum As Integer, _
    Optional ByVal lngLastUseID As Long, Optional ByVal lngShareUseID As Long, _
    Optional ByVal strBill As String, Optional strUseType As String = "", _
    Optional lngOutUseID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取张数够用并且指定票据在其可用范围内的领用ID
    '入参: bytKind      =   票种
    '      intNum       =   要打印的票据张数
    '      lngLastUseID =   上次使用的领用ID
    '      lngShareUseID=   本地参数指定的共用ID
    '      strBill      =   当前票据号，用于检查领用批次的票据范围
    '      strUseType-使用类别
    '      strUserName-操作员姓名
    '出参:lngOutUseID-返回领用ID
    '      >0   =   成功，可用的领用ID
    '      =0   =   失败
    '      -1   =   没有自用(用完或不够，或未领用),未设置共用
    '      -2   =   没有自用(用完或不够，或未领用),设置的共用已用完或不够
    '      -3   =   指定票据号不在当前所有可用领用批次的有效票据号范围内
    '      -4   =   指定批次的票据不够用
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-27 12:03:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strPre As String
    Dim blnTmp As Boolean, i As Integer, lngReturn As Long
    
    On Error GoTo errH
    '1.上次的领用批次是否可用并够用
    If lngLastUseID > 0 Then
        strSql = "" & _
        "   Select 前缀文本,开始号码,终止号码" & vbNewLine & _
        "   From 票据领用记录 " & _
        "   Where 票种=[1] And 剩余数量>=[2] And ID=[3]  " & _
        "           And (Nvl(使用类别,'LXH')=[4] Or  使用类别 Is NULL) "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "可用票据批次", bytKind, intNum, lngLastUseID, IIf(Trim(strUseType) = "", "LXH", strUseType))
        With rsTmp
            If .RecordCount > 0 Then    '目前的票据号可能和上次不同，所以需要检查范围
                If strBill = "" Then lngOutUseID = lngLastUseID: GetInvoiceGroupID = True: Exit Function '可能没有当前票据号
                blnTmp = False
                strPre = "" & !前缀文本
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then lngOutUseID = lngLastUseID: GetInvoiceGroupID = True: Exit Function
                
            ElseIf intNum > 1 Then  '不是确定领用批次调用时,当前票据号所在批次不够用
                lngOutUseID = -4: Exit Function
            End If
        End With
    End If
    
    '2.上次的领用批次不够用或不可用时,取自已领的并且自用的
    '  有多批，最近使用的优先,少的先用,先领先用
    strSql = "" & _
    "   Select ID, 前缀文本, 开始号码, 终止号码" & vbNewLine & _
    "   From 票据领用记录" & vbNewLine & _
    "   Where 票种 = [1] And 剩余数量 >= [2] And 领用人 = [3]  " & _
    "           And (Nvl(使用类别,'LXH')=[4] Or  使用类别 Is NULL ) " & _
    "           And 使用方式 = 1" & vbNewLine & _
    "   Order By Nvl(使用时间, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,使用类别 desc, 开始号码"
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "可用票据批次", bytKind, intNum, strUserName, IIf(strUseType = "", "LXH", strUseType))
    With rsTmp
        For i = 1 To .RecordCount
            If strBill = "" Then lngOutUseID = Val(Nvl(!ID)):   GetInvoiceGroupID = True: Exit Function '第一次使用时没有当前票据号
            blnTmp = False
            strPre = "" & !前缀文本
            If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                blnTmp = True
            End If
            If Not blnTmp Then lngOutUseID = Val(Nvl(!ID)):   GetInvoiceGroupID = True:  Exit Function
            .MoveNext
        Next
        lngReturn = IIf(.RecordCount > 0, -3, -1)
    End With
        
    '3.没有自用的,使用本地参数指定的共用批次
    If lngShareUseID > 0 Then
        strSql = "" & _
        "   Select 前缀文本,开始号码,终止号码" & vbNewLine & _
        "   From 票据领用记录  " & _
        "   Where 票种=[1] And 剩余数量>=[2] And ID=[3] " & _
        "   And (Nvl(使用类别,'LXH')=[4] Or  使用类别 Is NULL) "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "可用票据批次", bytKind, intNum, lngShareUseID, IIf(strUseType = "", "LXH", strUseType))
        With rsTmp
            If .RecordCount > 0 Then
                If strBill = "" Then lngOutUseID = lngShareUseID:   GetInvoiceGroupID = True:  Exit Function '第一次使用时没有当前票据号
                blnTmp = False
                strPre = "" & !前缀文本
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then lngOutUseID = lngShareUseID:   GetInvoiceGroupID = True:  Exit Function
            End If
            lngReturn = IIf(.RecordCount > 0, -3, -2)
        End With
    End If
    lngOutUseID = lngReturn
    GetInvoiceGroupID = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function zlGetInvoiceGroupID(ByVal lngModule As Long, _
    ByVal strUserName As String, ByVal bytBillType As gEM_BillType, ByVal strUserType As String, ByVal lngLastUserID As Long, _
    ByVal lngShareUseID As Long, ByRef lngOutUserID As Long, _
    Optional intNum As Integer = 1, Optional strCurInvoiceNo As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票的领用ID
    '入参: lngModule -模块号
    '    bytBillType -票种
    '    intNum -需要的发票数
    '    lngLastUserID -上次使用的领用ID
    '    lngShareUseID -共用票据ID
    '    strCurInvoiceNo -当前票号
    '    strUserType -使用类别
    '    strUserName-操作员姓名
    '出参:lngOutUserID-返回领用ID
    '     >0   =   成功，可用的领用ID
    '     =0   =   获取失败
    '     -1   =   没有自用(用完或不够，或未领用),未设置共用
    '     -2   =   没有自用(用完或不够，或未领用),设置的共用已用完或不够
    '     -3   =   指定票据号不在当前所有可用领用批次的有效票据号范围内
    '     -4   =   指定批次的票据不够用
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-27 11:58:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetInvoiceGroupID = GetInvoiceGroupID(strUserName, bytBillType, intNum, lngLastUserID, lngShareUseID, strCurInvoiceNo, strUserType, lngOutUserID)
End Function

Public Function zlGetNextBill(ByVal lngModule As Long, ByVal lng领用ID As Long, _
    ByRef strOutNextInvoiceNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据领用批次ID,获取下一个实际票据号
    '入参:lngModule-模块号
    '     lng领用ID-领用ID
    '出参:strOutNextInvoiceNo-一下张发票号
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-27 13:46:43
    '说明：1.当取不到范围内的有效票据时,返回空由用户输入
    '      2.排开已报损的号码
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsMain As ADODB.Recordset
    Dim rsDelete As ADODB.Recordset
    Dim strSql As String, strBill As String
    
    On Error GoTo errH
    strOutNextInvoiceNo = ""
    
    strSql = "Select 前缀文本,开始号码,终止号码,当前号码" & _
        " From 票据领用记录 Where 剩余数量>0 And ID=[1]"
    Set rsMain = gobjDatabase.OpenSQLRecord(strSql, "取一下票据号", lng领用ID)
    If rsMain.EOF Then Exit Function
    
    If IsNull(rsMain!当前号码) Then
        strBill = UCase(rsMain!开始号码)
    Else
        strBill = UCase(IncStr(rsMain!当前号码))
    End If
    strSql = "Select Upper(号码) as 号码 From 票据使用明细" & _
        " Where 号码||''>=[1] And 领用ID=[2]" & _
        " Order by 号码"
        
    Set rsDelete = gobjDatabase.OpenSQLRecord(strSql, "取一下票据号", strBill, lng领用ID)
    Do While True
        '检查范围
        If Left(strBill, Len("" & rsMain!前缀文本)) <> UCase("" & rsMain!前缀文本) Then
            Exit Function
        ElseIf Not (strBill >= UCase(rsMain!开始号码) And strBill <= UCase(rsMain!终止号码)) Then
            Exit Function
        End If
                
        '排开报损号
        rsDelete.Filter = "号码='" & UCase(strBill) & "'"
        If rsDelete.EOF Then Exit Do
        strBill = IncStr(strBill)
    Loop
    strOutNextInvoiceNo = strBill
    zlGetNextBill = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function zlStartFactUseType(ByVal bytBillType As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否使用了使用类别的
    '入参:bytBillType-票种
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-10 16:11:47
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    On Error GoTo errHandle
    strSql = "Select  1 as 存在 From 票据领用记录 where 票种=[1] and nvl(使用类别,'LXH')<>'LXH' and Rownum=1"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "检查票据是否启用了使用类别的", bytBillType)
    
    If rsTemp.EOF Then
        Set rsTemp = Nothing: Exit Function
    End If
    Set rsTemp = Nothing
    zlStartFactUseType = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-20 17:55:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDbUser
    If zlGetComLib = False Then Exit Function
    Call GetUserInfo
    Call InitVar
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCheckInvoiceOverplusEnough(ByVal bytKind As Byte, _
    ByVal intNum As Integer, Optional lng剩余数量 As Long, _
    Optional lng领用ID As Long = 0, Optional str使用类别 As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查票据的剩余数量是否充足
    '入参:bytKind-1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
    '     intNum-当前对比的数量(-1代表不提醒)
    '     lng领用ID-只检查当前的领用票据(32455)
    '     str使用类别-使用类别
    '出参:lng剩余数量-返回当前剩余数量
    '返回:充足返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-28 17:16:16
    '问题:26948
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    '-1代表不提醒
    If intNum = -1 Then zlCheckInvoiceOverplusEnough = True: Exit Function
    Err = 0: On Error GoTo ErrHand:
    
    lng剩余数量 = 0
    
    strSql = "" & _
        "   Select Sum(nvl(剩余数量,0)) as 剩余数量 " & vbNewLine & _
        "   From 票据领用记录" & vbNewLine & _
        "   Where 票种 = [1]  " & _
        "               And (nvl(使用类别,'LXH')=[4] or nvl(使用类别,'LXH')='LXH')  " & _
        "               And 领用人 = [2] And 使用方式 = 1 and nvl(剩余数量,0)>0" & vbNewLine & _
                    IIf(lng领用ID = 0, "", "             and ID=[3]") & _
        "   Union ALL " & _
        "   Select Sum(nvl(剩余数量,0)) as 剩余数量  " & _
        "   From 票据领用记录 A,人员表 B" & vbNewLine & _
          " Where A.票种=[1] And A.使用方式=2 And A.剩余数量>0 And A.领用人=B.姓名" & _
        "             And (nvl(A.使用类别,'LXH')=[4] or nvl(A.使用类别,'LXH')='LXH')  " & _
          "           And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
                       IIf(lng领用ID = 0, "", "             and A.ID=[3]") & _
          "  "
    strSql = "Select sum(剩余数量) as 剩余数量 From (" & strSql & ")"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "可用票据批次", bytKind, UserInfo.姓名, lng领用ID, str使用类别)
    lng剩余数量 = Val(Nvl(rsTemp!剩余数量))
    zlCheckInvoiceOverplusEnough = lng剩余数量 > intNum
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then Resume
End Function

Public Function zlCheckBillNOIsPrintInvoice(ByVal byt票种 As gEM_BillType, ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定的单据是否已经存在打印票据的情况
    '入参:byt票种-1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
    '     strNOs-指定要重打的单据号，带引号，可能是多个单据号，为"'AAA','BBB',..."的形式
    '返回:存在票据返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-30 10:29:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSql = "" & _
    "   Select /*+ rule */ Max(A.ID) as ID " & _
    "   From 票据打印内容 A,Table( f_Str2list([1])) J " & _
    "   Where A.数据性质=[2]   And A.NO=J.Column_Value"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "mdlOutExse", Replace(strNos, "'", ""), byt票种)
    If rsTemp.RecordCount <> 0 Then
        zlCheckBillNOIsPrintInvoice = Val(Nvl(rsTemp!ID)) <> 0
    End If
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Class_Initialize()
    glngInstanceCount = glngInstanceCount + 1
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub
