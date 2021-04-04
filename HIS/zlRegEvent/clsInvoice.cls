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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*********************************************************************************************************************************************
'功能:票据控制

'编制:LNC
'日期:2020-03-19 10:43:04
'*********************************************************************************************************************************************
Private mfrmMain As Object
Private mblnSharedInvoice As Boolean
Private mblnStartFactUseType As Boolean
Private mbln严格控制 As Boolean
Private mblnHisMgrInvoice As Boolean 'his管理票据，启用电子票据时有效
Private mblnPrint As Boolean
Private mbyt场合 As Byte
Private mbytInvoiceLen As Byte
Private mbytInvoiceType As Byte '发票打印方式
Private mbytInvoiceFormat As Byte '发票打印格式
Private mlngModule As Long
Private mlng领用ID As Long
Private mlng共用批次ID As Long
Private mstrUseType As String
Private mstr场合 As String
Private mstrInvoice As String

Public Property Let strInvoice(ByVal varValue As String)
    mstrInvoice = varValue
End Property

Public Property Get strInvoice() As String
    strInvoice = mstrInvoice
End Property

Public Property Get strUseType() As String
    strUseType = mstrUseType
End Property

Public Property Let blnPrint(ByVal varValue As Boolean)
    mblnPrint = varValue
End Property

Public Property Get blnPrint() As Boolean
    blnPrint = mblnPrint
End Property

Public Property Let blnHisMgrInvoice(ByVal varValue As Boolean)
    mblnHisMgrInvoice = varValue
End Property

Public Property Get blnHisMgrInvoice() As Boolean
    blnHisMgrInvoice = mblnHisMgrInvoice
End Property

Public Property Get bytPrintType() As Byte
    bytPrintType = mbytInvoiceType
End Property

Public Property Get bytPrintFormat() As Byte
    bytPrintFormat = mbytInvoiceFormat
End Property

Public Property Get blnSharedInvoice() As Boolean
    blnSharedInvoice = mblnSharedInvoice
End Property

Public Property Get blnStartFactUseType() As Boolean
    blnStartFactUseType = mblnStartFactUseType
End Property

Public Property Get bln严格控制() As Boolean
    bln严格控制 = mbln严格控制
End Property

Public Property Get lng共用批次ID() As Long
    lng共用批次ID = mlng共用批次ID
End Property

Public Property Get lng领用ID() As Long
    lng领用ID = mlng领用ID
End Property


Public Function zlInitialize(ByVal frmMain As Object, Optional ByVal byt场合 As Byte = 4, Optional ByVal lngModule As Long = 1111) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化票据对象
    '入参:
    '    byt场合：4-挂号;5-发卡
    '编制:刘兴洪
    '日期:2020-03-03 11:17:34
    '说明:
    '    1.使用本部件前,必须先调用本接口进行初始化
    '    2.初始化接口,在HIS进入模块时调用(例如：进入收费管理界面)
    '    3.模块号：1111=门诊挂号,1121=门诊收费,1107=医疗卡发放管理;
    '调用者:
    '   挂号?挂号票据设置?
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain: mlngModule = lngModule
    mbyt场合 = byt场合
    Select Case byt场合
        Case 4
            mstr场合 = "挂号"
        Case 5
            mstr场合 = "收费"
    End Select
    
    If zlInitPara(byt场合) = False Then Exit Function
    zlInitialize = True
End Function

Public Function zlTerminate() As Boolean
    Set mfrmMain = Nothing
End Function

Private Function zlInitPara(ByVal byt场合 As Byte) As Boolean
    Dim strValue As String
    
    '挂号发卡也使用挂号打印方式控制
    mbytInvoiceType = Val(zlDatabase.GetPara("挂号发票打印方式", glngSys, mlngModule))
    If byt场合 = 4 Then
        '挂号使用收费票据
        mblnSharedInvoice = zlDatabase.GetPara("挂号共用收费票据", glngSys, 1121) = "1"
        
    ElseIf byt场合 = 5 Then
        '医疗卡使用收费票据
        mblnSharedInvoice = zlDatabase.GetPara("卡费使用门诊收费医疗收据", glngSys, 1107) = "1"
        If Not mblnSharedInvoice Then zlInitPara = True: Exit Function
        
        mbytInvoiceFormat = Val(Split(zlDatabase.GetPara("医疗卡收据格式", glngSys, 1107) & "|", "|")(0))
    End If
    
    '本地共用挂号批次ID
    If mblnSharedInvoice Then
        mlng共用批次ID = Val(zlDatabase.GetPara("共用收费票据批次", glngSys, 1121, ""))
    Else
        mlng共用批次ID = Val(zlDatabase.GetPara("共用挂号票据批次", glngSys, mlngModule, ""))
    End If
    If mlng共用批次ID > 0 Then
        If Not ExistShareBill(mlng共用批次ID, IIf(mblnSharedInvoice, 1, 4)) Then
            Call SaveCurInvoice("0")
            mlng共用批次ID = 0
        End If
    End If
    '是否区分收费票据使用类别
    If mblnSharedInvoice Then
        mblnStartFactUseType = zlStartFactUseType("1")
    End If
    
    '票据是否严格控制
    strValue = zlDatabase.GetPara(24, glngSys, , "00000")
    mbln严格控制 = (Mid(strValue, IIf(mblnSharedInvoice, 1, 4), 1) = "1")
    '票号长度
    strValue = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    mbytInvoiceLen = Val(Split(strValue, "|")(IIf(gblnSharedInvoice, 0, 3)))
    
    zlInitPara = True
End Function

Public Function GetUseType(ByVal lng病人id As Long, ByVal intInsure As Integer) As String
    '功能:获取收费票据的使用类别
    Dim strUseType As String
    Dim varData As Variant, varTemp As Variant
    mstrUseType = "": mlng共用批次ID = 0: mlng领用ID = 0
    If mblnStartFactUseType Then
        mstrUseType = zl_GetInvoiceUserType(lng病人id, 0, intInsure)
        mlng共用批次ID = zl_GetInvoiceShareID(1121, strUseType)
        If mbytInvoiceType <> 0 Then
            Call zlGetInvoiceGroupUseID(mlng领用ID)
        End If
    End If
    GetUseType = mstrUseType
End Function

Public Function SaveCurInvoice(ByVal strInvoice As String) As Boolean
    '功能：保存当前使用的票据号
    If mbln严格控制 Then SaveCurInvoice = True: Exit Function
    If mbyt场合 = 4 Then
        If mblnSharedInvoice Then
            zlDatabase.SetPara "当前收费票据号", strInvoice, glngSys, 1121
        Else
            zlDatabase.SetPara "当前挂号票据号", strInvoice, glngSys, mlngModule
        End If
    Else
        zlDatabase.SetPara "当前收费票据号", strInvoice, glngSys, 1121
    End If
    SaveCurInvoice = True
End Function

Public Function zlGetInvoiceGroupUseID(ByRef lng领用ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的领用ID
    '入参:lng领用ID-领用id
    '       intNum-页数
    '       strInvoiceNO-输入的发票号
    '出参:lng领用ID-领用ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-11-19 16:32:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mbyt场合 <> 4 Then Exit Function
    mlng领用ID = GetInvoiceGroupID(IIf(mblnSharedInvoice, 1, 4), intNum, lng领用ID, mlng共用批次ID, strInvoiceNO, IIf(mblnStartFactUseType, mstrUseType, ""))
    If mlng领用ID <= 0 Then
        Select Case mlng领用ID
            Case 0 '操作失败
            Case -1
                If Trim(mstrUseType) = "" Then
                    MsgBox "你没有自用和共用的收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "你没有自用和共用的『" & mstrUseType & "』收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
            Case -2
                If Trim(mstrUseType) = "" Then
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "本地的共用票据的『" & mstrUseType & "』收费票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
            Case -3
                MsgBox "当前票据号码不在可用领用批次的有效票据号范围内,请重新输入！", vbInformation, gstrSysName
                On Error Resume Next
                zlControl.ControlSetFocus mfrmMain.txtFact
                Exit Function
        End Select
    End If
    lng领用ID = mlng领用ID
    zlGetInvoiceGroupUseID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCheckUsedBill(ByVal lng领用ID As Long, Optional ByVal strInvoice As String) As Long
    '获取领用ID，如果是使用收费票据，在调用前必须要先调用GetUseType()
    If mbyt场合 = 4 Then
        mlng领用ID = CheckUsedBill(IIf(mblnSharedInvoice, 1, 4), IIf(lng领用ID > 0, lng领用ID, mlng共用批次ID), strInvoice, mstrUseType)
    ElseIf mbyt场合 = 5 Then
        If Not mblnSharedInvoice Then Exit Function
        mlng领用ID = CheckUsedBill(1, IIf(lng领用ID > 0, lng领用ID, mlng共用批次ID), strInvoice, mstrUseType)
    Else
        mlng领用ID = 0
    End If
    zlCheckUsedBill = mlng领用ID
End Function

Public Function zlCheckBillRepeat(ByVal lng领用ID As Long, ByVal strInvoice As String) As Boolean
    '检查并发性
    If mbyt场合 = 4 Then
        zlCheckBillRepeat = CheckBillRepeat(lng领用ID, IIf(mblnSharedInvoice, 1, 4), strInvoice)
    ElseIf mbyt场合 = 5 Then
        If Not mblnSharedInvoice Then Exit Function
        zlCheckBillRepeat = CheckBillRepeat(lng领用ID, 1, strInvoice)
    End If
End Function

Public Function zlCheckInvoiceIsValid(ByRef lng领用ID As Long, Optional ByRef strInvoice As String, Optional ByVal blnAutoInput As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查票据有效性
    ' 入参 : strInvoice：实际票号
    '        blnAutoInput:票号是否自动获取的
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/8 19:25
    '---------------------------------------------------------------------------------------
    If Not mblnPrint Then zlCheckInvoiceIsValid = True: Exit Function
    
    If lng领用ID = 0 Then lng领用ID = mlng领用ID
    If mbln严格控制 Then
        If strInvoice = "" Then
            If mbyt场合 = 4 Then
                MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
                Exit Function
            Else
                '自动产生发卡收费票据
                strInvoice = zlGetNextInvoice(lng领用ID)
            End If
        End If
InvoiceHandle:
        lng领用ID = zlCheckUsedBill(lng领用ID, strInvoice)
        If lng领用ID <= 0 Then
            Select Case lng领用ID
            Case 0    '操作失败
            Case -1
                MsgBox "你没有自用和共用的" & mstr场合 & "票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            Case -2
                MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            Case -3
                MsgBox "票据号码不在当前有效领用范围内,请重新输入！", vbInformation, gstrSysName
            End Select
            Exit Function
        End If
        
        '并发操作检查,票号是否已用
        If zlCheckBillRepeat(lng领用ID, strInvoice) Then
            If Not blnAutoInput Then
                MsgBox "票据号""" & strInvoice & """已经被使用，请重新输入。", vbInformation, gstrSysName
                Exit Function
            Else
                strInvoice = zlGetNextInvoice(lng领用ID)
                If strInvoice = "" Then Exit Function
                
                MsgBox "当前票据号已经被使用，已重新获取票据号:" & strInvoice, vbInformation, gstrSysName
                GoTo InvoiceHandle
            End If
        End If
    Else
        If Len(strInvoice) <> mbytInvoiceLen And strInvoice <> "" Then
            MsgBox "票据号码长度应该为 " & mbytInvoiceLen & " 位！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    mlng领用ID = lng领用ID
    zlCheckInvoiceIsValid = True
End Function

Public Function zlGetNextInvoice(ByRef lng领用ID As Long) As String
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取下一个有效发票号
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/8 19:37
    '---------------------------------------------------------------------------------------
    Dim strCurInvoice As String
    
    If mbln严格控制 Then
        lng领用ID = zlCheckUsedBill(lng领用ID)
        If lng领用ID <= 0 Then
            Select Case lng领用ID
                Case 0 '操作失败
                Case -1
                    MsgBox "你没有自用和共用的" & mstr场合 & "票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End Select
            Exit Function
        End If
        '严格：取下一个号码
        zlGetNextInvoice = GetNextBill(lng领用ID)
    Else
        '松散：取下一个号码
        If mblnSharedInvoice Then
            strCurInvoice = zlDatabase.GetPara("当前收费票据号", glngSys, 1121)
        Else
            strCurInvoice = zlDatabase.GetPara("当前挂号票据号", glngSys, mlngModule)
        End If
        zlGetNextInvoice = zlStr.Increase(strCurInvoice)
    End If
End Function
