VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBalanceAllCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'保持属性值的局部变量
Private mvarstrAllTime As String '局部复制：病人所有未结帐住院次数
Private mvarstrUnAuditTime As String '局部复制：病人所有未审核住院次数
Private mvarstrAllDeptIDs As String '局部复制：病人所有未结帐科室
Private mvarstrAllItem As String '局部复制：病人所有未结收据费目
Private mvarstrAllClass As String '局部复制：病人所有未结费用类型
Private mvarstrAllChargeType As String '局部复制：病人所有未结的收费类别
Private mvarstrAllOwnerFeeType As String '局部复制：病人所有未结的自费的收费类别
Private mvarstrAllDiag As String '局部复制：所有诊断
Private mvarstrAllBabys As String '局部复制：所有婴儿
Private mvarMinDate As Date '局部复制
Private mvarMaxDate As Date '局部复制
Private mvarrsALLTime As ADODB.Recordset '局部复制
Private mvarstrAllFullTims As String '局部复制
Public Property Let strAllFullTims(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.strAllFullTims = 5
    mvarstrAllFullTims = vData
End Property


Public Property Get strAllFullTims() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.strAllFullTims
    strAllFullTims = mvarstrAllFullTims
End Property




Public Property Set rsAllTime(ByVal vData As ADODB.Recordset)
    '向属性指派对象时使用，位于 Set 语句的左边。
    'Syntax: Set x.rsALLTime = Form1
    Set mvarrsALLTime = vData
End Property


Public Property Get rsAllTime() As ADODB.Recordset
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.rsALLTime
    Set rsAllTime = mvarrsALLTime
End Property
Public Property Let MaxDate(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.MaxDate = 5
    mvarMaxDate = vData
End Property


Public Property Get MaxDate() As Date
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.MaxDate
    MaxDate = mvarMaxDate
End Property



Public Property Let MinDate(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.MinDate = 5
    mvarMinDate = vData
End Property


Public Property Get MinDate() As Date
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.MinDate
    MinDate = mvarMinDate
End Property



Public Property Let strAllBabys(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.strAllBabys = 5
    mvarstrAllBabys = vData
End Property


Public Property Get strAllBabys() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.strAllBabys
    strAllBabys = mvarstrAllBabys
End Property



Public Property Let strAllOwnerFeeType(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.strAllOwnerFeeType = 5
    mvarstrAllOwnerFeeType = vData
End Property


Public Property Get strAllOwnerFeeType() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.strAllOwnerFeeType
    strAllOwnerFeeType = mvarstrAllOwnerFeeType
End Property



Public Property Let strAllChargeType(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.strAllChargeType = 5
    mvarstrAllChargeType = vData
End Property


Public Property Get strAllChargeType() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.strAllChargeType
    strAllChargeType = mvarstrAllChargeType
End Property


Public Property Let strAllDiag(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.strAllChargeType = 5
    mvarstrAllDiag = vData
End Property


Public Property Get strAllDiag() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.strAllChargeType
    strAllDiag = mvarstrAllDiag
End Property


Public Property Let strAllClass(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.strAllClass = 5
    mvarstrAllClass = vData
End Property


Public Property Get strAllClass() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.strAllClass
    strAllClass = mvarstrAllClass
End Property



Public Property Let strAllItem(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.strAllItem = 5
    mvarstrAllItem = vData
End Property


Public Property Get strAllItem() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.strAllItem
    strAllItem = mvarstrAllItem
End Property



Public Property Let strAllDeptIDs(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.strAllDeptIDs = 5
    mvarstrAllDeptIDs = vData
End Property


Public Property Get strAllDeptIDs() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.strAllDeptIDs
    strAllDeptIDs = mvarstrAllDeptIDs
End Property



Public Property Let strUnAuditTime(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.strUnAuditTime = 5
    mvarstrUnAuditTime = vData
End Property


Public Property Get strUnAuditTime() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.strUnAuditTime
    strUnAuditTime = mvarstrUnAuditTime
End Property



Public Property Let strAllTime(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.strAllTime = 5
    mvarstrAllTime = vData
End Property


Public Property Get strAllTime() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.strAllTime
    strAllTime = mvarstrAllTime
End Property

Public Function zlGetTimeRecordFromTimeString(ByVal lng病人ID As Long, ByVal strAllTimes As String, _
    ByRef rsTimes As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据住院次数串，获取每次信息信息
    '入参:
    '出参:rsTimes-住院信息集(病人ID,主页ID,险类,险类名称)
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-12 14:37:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    Set rsTimes = Nothing
    
    strSQL = "" & _
    "   Select A.病人ID,A.主页ID,Nvl(A.险类,0) As 险类,B.名称 as 保险名称" & _
    "   From 病案主页 A,保险类别 B " & _
    "   Where A.险类=B.序号(+) And A.病人ID = [1]  " & _
    "       And A.主页ID In (Select Column_Value From Table(f_str2list([2]))) " & _
    "   Order By A.病人ID,A.主页ID Desc"
    
    Set rsTimes = zlDatabase.OpenSQLRecord(strSQL, "获取指定次数的住院信息", lng病人ID, strAllTimes)
    
    zlGetTimeRecordFromTimeString = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetNumsFromZyNo(ByVal str住院号 As String, Optional ByRef lng病人ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据住院号来获取住院次数
    '入参:lng住院号
    '出参:lng病人ID-返回病人ID
    '返回:住院次数,多次用逗号分离
    '编制:刘兴洪
    '日期:2017-11-12 10:40:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strTimeNums As String
     
    lng病人ID = 0
    strSQL = "Select 病人ID,主页ID From 病案主页 Where 住院号 = [1] And Nvl(病人性质,0) = 0 Order By 主页ID DESC"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据住院号来获取所涉及的住院次数", str住院号)
    strTimeNums = ""
    Do While Not rsTemp.EOF
        lng病人ID = Val(Nvl(rsTemp!病人ID))
        strTimeNums = strTimeNums & "," & Val(Nvl(rsTemp!主页ID))
        rsTemp.MoveNext
    Loop
    If strTimeNums <> "" Then strTimeNums = Mid(strTimeNums, 2)
    zlGetNumsFromZyNo = strTimeNums
End Function
Public Function GetPatiFeeDateRang(ByVal lng病人ID As Long, str主页Ids As String, ByRef strStartDate As String, ByRef strEndDate As String, ByRef bln登记时间 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人的费用时间范围
    '入参:lng病人ID
    '     str主页IDs-主页ID,多个用逗号分隔
    '出参:strStartDate-开如时间
    '     strEndDate-结束时间
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2017-11-26 10:49:03
    '说明:费用取数规则如下:
    '    一、开始时间
    '       1.如果入院日期小于费用时间，则取入院时间
    '       2.如果入院日期大于费用时间，则取费用时间
    '    二、结束时间
    '       1.如果出院日期小于费用时间，则取费用时间
    '       2.如果出院日期大于费用时间，则取出院时间
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFeeMinDate As String, strFeeMaxDate As String
    Dim str入院日期 As String, str出院日期 As String
    Dim dtDate As Date, dtTemp As Date
    
    On Error GoTo errHandle
    strStartDate = "": strEndDate = ""
    
    '取费用的最小或最大时间
    Call GetPatiFeeMinAndMaxDate(lng病人ID, str主页Ids, strFeeMinDate, strFeeMaxDate, bln登记时间)
    Call GetPatiInOutDate(lng病人ID, str主页Ids, str入院日期, str出院日期)
    
    
    If str入院日期 <> "" Then
        dtDate = CDate(str入院日期)
        If strFeeMinDate <> "" Then
            dtTemp = CDate(strFeeMinDate)
            If dtDate > dtTemp Then dtDate = dtTemp
        End If
    Else
        dtDate = zlDatabase.Currentdate
    End If
    strStartDate = Format(dtDate, "yyyy-mm-dd HH:MM:SS")
    
    If str出院日期 <> "" Then
        dtDate = CDate(str出院日期)
        If strFeeMaxDate <> "" Then
            dtTemp = CDate(strFeeMaxDate)
            If dtDate < dtTemp Then dtDate = dtTemp
        End If
        strEndDate = Format(dtDate, "yyyy-mm-dd HH:MM:SS")
    Else
        '未出院
        If strFeeMaxDate <> "" Then
            dtDate = CDate(strFeeMaxDate)
        Else
            dtDate = zlDatabase.Currentdate
        End If
        strEndDate = Format(dtDate, "yyyy-mm-dd HH:MM:SS")
    End If
    If strStartDate > strEndDate Then
        strEndDate = strStartDate
    End If
    GetPatiFeeDateRang = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetPatiInOutDate(lng病人ID As Long, str主页Ids As String, ByRef str入院时间 As String, ByRef str出院日期 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人某次住院的入院或出院时间
    '入参:
    '出参:str入院时间-格式:yyyy-mm-dd HH:MM:SS：主页ID为多次时,返回最小的入院时间
    '     str出院日期-格式:yyyy-mm-dd HH:MM:SS：主页ID为多次时,返回最大的入院时间,未出院时，出院日期为""
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2017-11-26 11:09:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strWhere As String, lng主页ID As Long
    
    On Error GoTo errHandle
    
    If InStr(str主页Ids, ",") > 0 Then
        strWhere = " And 主页ID in (Select Column_Value From Table(f_num2list([2]))) "
    Else
        strWhere = " And 主页ID= [3] ": lng主页ID = Val(str主页Ids)
    End If
    
    strSQL = "Select min(入院日期) as 入院日期,max(出院日期) as 出院日期 From 病案主页 Where 病人ID=[1]  " & strWhere
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取病人的入出院日期", lng病人ID, str主页Ids, lng主页ID)
    
    str入院时间 = Format(rsTmp!入院日期, "yyyy-mm-dd HH:MM:SS")
    str出院日期 = Format(rsTmp!出院日期, "yyyy-mm-dd HH:MM:SS")
 
    GetPatiInOutDate = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetPatiFeeMinAndMaxDate(ByVal lng病人ID As Long, ByVal str主页Ids As String, ByRef strMinDate As String, ByRef strMaxDate As String, Optional bln登记时间 As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定病人的登记时间
    '入参:lng病人ID
    '     str主页IDs-主页ID,多个用逗号分隔
    '出参:strMinDate-格式:yyyy-mm-dd HH:MM:SS
    '     strMaxDate-格式:yyyy-mm-dd HH:MM:SS
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2013-10-22 17:16:47
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strFields As String, strWhere As String, lng主页ID As Long
    
    On Error GoTo errHandle
    strFields = IIf(bln登记时间, "登记时间", "发生时间")
    strFields = "Min(" & strFields & ") as 最小时间,Max(" & strFields & ") as 最大时间 "
    If InStr(str主页Ids, ",") > 0 Then
        strWhere = " And 主页ID in (Select Column_Value From Table(f_num2list([2]))) "
    Else
        strWhere = " And 主页ID= [3] ": lng主页ID = Val(str主页Ids)
    End If
    strSQL = " Select " & strFields & " From 住院费用记录 where mod(记录性质,10) in (2,3) And 病人ID=[1] " & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取费用的最小和最大时间", lng病人ID, str主页Ids, lng主页ID)
    
    strMinDate = Format(rsTemp!最小时间, "yyyy-mm-dd HH:MM:SS")
    strMaxDate = Format(rsTemp!最大时间, "yyyy-mm-dd HH:MM:SS")
    GetPatiFeeMinAndMaxDate = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



