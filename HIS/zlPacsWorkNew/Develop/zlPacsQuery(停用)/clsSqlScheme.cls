VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSqlScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjXml As Object
Private mstrXmlVersion As String

Private mlngSchemeID As Long        '方案ID
Private mstrSchemeName As String    '方案名称
Private mstrDescript As String      '方案描述
Private mdblVer As Double           '方案版本
Private mblnUseCard As Boolean      '是否启用刷卡
Private mblnUseGroup As Boolean     '是否启用分组
Private mblnShowHistory As Boolean  '是否显示历史检查
Private mblnUseFuncFollow As Boolean '是否启用功能跟随
Private mlngPageRecord As Long      '分页记录大小
Private mlngDateRange As Long       '日期选择范围
Private mstrStore As String         '恢复段

Private mstrQuery As String         '查询语句
Private mstrDetail As String        '明细语句

Private mobjSerachCfg() As clsScSerachCfg   'serach配置
Private mobjFilterCfg() As clsScFilterCfg   'filter配置
Private mobjShowCfg() As clsScShowCfg       'Show配置

Private aryCfgConvertCol() As String
Private aryCfgConvertData() As String

Private Const SCHEME_MODULE As String = "<scheme [scheme]>" & vbCrLf & _
                                "        <query>" & vbCrLf & _
                                "            [query]" & vbCrLf & _
                                "        </query>" & vbCrLf & _
                                "        <detail>" & vbCrLf & _
                                "            [detail]" & vbCrLf & _
                                "        </detail>" & vbCrLf & _
                                "        <serachcfg>" & vbCrLf & _
                                "[serachcfg]" & vbCrLf & _
                                "        </serachcfg>" & vbCrLf & _
                                "        <filtercfg>" & vbCrLf & _
                                "[filtercfg]" & vbCrLf & _
                                "        </filtercfg>" & vbCrLf & _
                                "        <showcfg>" & vbCrLf & _
                                "[showcfg]" & vbCrLf & _
                                "        </showcfg>" & vbCrLf & _
                                "        <store>" & vbCrLf & _
                                "[store]" & vbCrLf & _
                                "        </store>" & vbCrLf & _
                                "</scheme>"
                                
Private Const NODETYPE_ELEMENT As Long = 1
Private Const NODETYPE_ATTRIBUTE As Long = 2
Private Const NODETYPE_TEXT As Long = 3


'SchemeID
Property Get SchemeId() As Long
    SchemeId = mlngSchemeID
End Property

Property Let SchemeId(ByVal value As Long)
    mlngSchemeID = value
End Property

'SchemeName
Property Get SchemeName() As String
    SchemeName = mstrSchemeName
End Property

Property Let SchemeName(ByVal value As String)
    mstrSchemeName = value
End Property


'ver
Property Get Ver() As Double
    Ver = mdblVer
End Property

Property Let Ver(ByVal value As Double)
    mdblVer = value
End Property


'Descript
Property Get Descript() As String
    Descript = mstrDescript
End Property

Property Let Descript(ByVal value As String)
    mstrDescript = value
End Property


'UseCard
Property Get UseCard() As Boolean
    UseCard = mblnUseCard
End Property

Property Let UseCard(ByVal value As Boolean)
    mblnUseCard = value
End Property


'UseGroup
Property Get UseGroup() As Boolean
    UseGroup = mblnUseGroup
End Property

Property Let UseGroup(ByVal value As Boolean)
    mblnUseGroup = value
End Property

'ShowHistory
Property Get ShowHistory() As Boolean
    ShowHistory = mblnShowHistory
End Property

Property Let ShowHistory(ByVal value As Boolean)
    mblnShowHistory = value
End Property

'UseFuncFollow
Property Get UseFuncFollow() As Boolean
    UseFuncFollow = mblnUseFuncFollow
End Property

Property Let UseFuncFollow(ByVal value As Boolean)
    mblnUseFuncFollow = value
End Property

'PageRecord
Property Get PageRecord() As Long
    PageRecord = mlngPageRecord
End Property

Property Let PageRecord(ByVal value As Long)
    mlngPageRecord = value
End Property

'daterange
Property Get DateRange() As Long
    DateRange = mlngDateRange
End Property

Property Let DataRange(ByVal value As Long)
    mlngDateRange = value
End Property

'store
Property Get Store() As String
    Store = mstrStore
End Property

Property Let Store(ByVal value As String)
    mstrStore = value
End Property

'query
Property Get Query() As String
    Query = mstrQuery
End Property

Property Let Query(ByVal value As String)
    mstrQuery = value
End Property

'detail
Property Get Detail() As String
    Detail = mstrDetail
End Property

Property Let Detail(ByVal value As String)
    mstrDetail = Detail
End Property

'serach
Property Get SerachCfgCount()
    SerachCfgCount = UBound(mobjSerachCfg)
End Property

'filter
Property Get FilterCfgCount()
    FilterCfgCount = UBound(mobjFilterCfg)
End Property

'show
Property Get ShowCfgCount()
    ShowCfgCount = UBound(mobjShowCfg)
End Property


'serach
Property Get SerachCfg(ByVal lngIndex As Long) As clsScSerachCfg
    Set SerachCfg = mobjSerachCfg(lngIndex)
End Property

'filter
Property Get FilterCfg(ByVal lngIndex As Long) As clsScFilterCfg
    Set FilterCfg = mobjFilterCfg(lngIndex)
End Property

'show
Property Get ShowCfg(ByVal lngIndex As Long) As clsScShowCfg
    Set ShowCfg = mobjShowCfg(lngIndex)
End Property

'config convert col
Property Get ConvertColName(ByVal lngIndex As Long) As String
    ConvertColName = aryCfgConvertCol(lngIndex)
End Property

Property Get ConvertColNames() As String()
    ConvertColNames = aryCfgConvertCol
End Property

'config convert data
Property Get ConvertColData(ByVal lngIndex As Long) As String
    ConvertColData = aryCfgConvertData(lngIndex)
End Property

Property Get ConvertColDatas() As String()
    ConvertColDatas = aryCfgConvertData
End Property

Public Function RefreshConvertColCount(rsSourceData As ADODB.Recordset) As Long
    Dim lngCount As Long
    
    lngCount = UBound(aryCfgConvertCol)
    
    If lngCount <= 0 Then
        lngCount = LoadConvertConfig(rsSourceData)
    End If
    
    RefreshConvertColCount = lngCount
End Function


Private Function LoadConvertConfig(rsSourceData As ADODB.Recordset) As Long
    Dim i           As Long
    Dim lngUbound   As Long
    Dim objShowCfg  As clsScShowCfg
    
    ReDim aryCfgConvertCol(0)
    ReDim aryCfgConvertData(0)

    LoadConvertConfig = 0
    
    For i = 1 To ShowCfgCount
        Set objShowCfg = mobjShowCfg(i)
        
        '配置了数据转换，且该转换字段在数据集中存在，则需要进行转换处理
        If objShowCfg.DataConvert <> "" Then
            If HasField(rsSourceData, objShowCfg.Name) Then
                '进行数据转换
                lngUbound = UBound(aryCfgConvertCol) + 1
                
                ReDim Preserve aryCfgConvertCol(lngUbound)
                ReDim Preserve aryCfgConvertData(lngUbound)
                
                aryCfgConvertCol(lngUbound) = objShowCfg.Name
                aryCfgConvertData(lngUbound) = ";" & objShowCfg.DataConvert & ";"
            End If
        End If
    Next i
    
    LoadConvertConfig = UBound(aryCfgConvertCol)
End Function


Public Function GetScheme() As String
'获取方案内容串
    Dim i As Long
    Dim j As Long
    Dim strResult As String
    Dim strCfgs As String
    Dim objRowRelation As clsScRowRelation
    
    strResult = SCHEME_MODULE
    strCfgs = "id=""" & mlngSchemeID & """ name=""" & mstrSchemeName & _
                """ usecard=""" & mblnUseCard & """ showhistory=""" & mblnShowHistory & """ usegroup= """ & mblnUseGroup & """ usefuncfollow=""" & mblnUseFuncFollow & _
                """ pagerecord=""" & mlngPageRecord & """ daterange=""" & mlngDateRange & """ ver=""" & mdblVer & """ descript=""" & mstrDescript & """"
    
    strResult = Replace(strResult, "[scheme]", strCfgs)
    strResult = Replace(strResult, "[query]", mstrQuery)
    strResult = Replace(strResult, "[detail]", mstrDetail)
    
    '写入serach
    strCfgs = ""
    For i = 1 To SerachCfgCount
        If Len(strCfgs) > 0 Then strCfgs = strCfgs & vbCrLf
        
        '""" sourcetype=""" & mobjSerachCfg(i).SourceType
        strCfgs = strCfgs & "    <item name=""" & mobjSerachCfg(i).Name & _
                                                """ inputtype=""" & mobjSerachCfg(i).InputType & _
                                                """ controltype=""" & mobjSerachCfg(i).ControlType & _
                                                """ likeway=""" & mobjSerachCfg(i).LikeWay & _
                                                """ default=""" & mobjSerachCfg(i).Default & _
                                                """>" & FormatXml(mobjSerachCfg(i).DataFrom) & "</item>"
    Next i
    strResult = Replace(strResult, "[serachcfg]", strCfgs)
    
    '写入filter
    strCfgs = ""
    For i = 1 To FilterCfgCount
        If Len(strCfgs) > 0 Then strCfgs = strCfgs & vbCrLf
        
        '""" filtertype=""" & mobjFilterCfg(i).FilterType
        '""" sourcetype=""" & mobjFilterCfg(i).SourceType
        strCfgs = strCfgs & "    <item name=""" & mobjFilterCfg(i).Name & _
                                                """ selectway=""" & mobjFilterCfg(i).SelectWay & _
                                                """>" & FormatXml(mobjFilterCfg(i).DataFrom) & _
                                                IIf(mobjFilterCfg(i).CustomScript <> "", _
                                                    "<cusscript><![CDATA[" & mobjFilterCfg(i).CustomScript & "]]></cusscript>", _
                                                    "") & "</item>"
    Next i
    strResult = Replace(strResult, "[filtercfg]", strCfgs)
    
    
    '写入Show
    strCfgs = ""
    For i = 1 To ShowCfgCount
        If Len(strCfgs) > 0 Then strCfgs = strCfgs & vbCrLf
        
        strCfgs = strCfgs & "    <item name=""" & mobjShowCfg(i).Name & _
                                                """ icon=""" & mobjShowCfg(i).Icon & _
                                                """ hiddencol=""" & mobjShowCfg(i).HiddenCol & _
                                                """ hiddentitle=""" & mobjShowCfg(i).HiddenTitle & _
                                                """ hiddendata=""" & mobjShowCfg(i).HiddenData & _
                                                """ istotal=""" & mobjShowCfg(i).IsTotal & _
                                                """ uselistlocate=""" & mobjShowCfg(i).UseListLocate & _
                                                """ sortcontrastcol=""" & mobjShowCfg(i).SortContrastCol & _
                                                """>" & FormatXml(mobjShowCfg(i).DataConvert)
        If mobjShowCfg(i).RowRelationCount > 0 Then
            '写入rowleation
            For j = 1 To mobjShowCfg(i).RowRelationCount
                Set objRowRelation = mobjShowCfg(i).RowRelation(j)
                
                strCfgs = strCfgs & vbCrLf & "                <rowrelation tiggerdata=""" & objRowRelation.TiggerData & _
                                                """ icon=""" & objRowRelation.Icon & _
                                                """ iconperformcol=""" & objRowRelation.IconPerformCol & _
                                                """ isstateicon=""" & objRowRelation.IsStateIcon & _
                                                """ rowfc=""" & objRowRelation.RowFontColor & _
                                                """ rowbc=""" & objRowRelation.RowBackColor & _
                                                """ cellfc=""" & objRowRelation.CellFontColor & _
                                                """ cellbc=""" & objRowRelation.CellBackColor & _
                                                """ colorperformcol=""" & objRowRelation.ColorPerformCol & _
                                                """ flickertimeout=""" & objRowRelation.FlickerTimeOut & _
                                                """ timeoutrefercol=""" & objRowRelation.TimeOutReferCol & """ />"
            Next j
        End If
                                                
        strCfgs = strCfgs & vbCrLf & "</item>"
    Next i
    strResult = Replace(strResult, "[showcfg]", strCfgs)
    strResult = Replace(strResult, "[store]", "<![CDATA[" & mstrStore & "]]>")
    
    GetScheme = strResult
End Function

Private Function FormatXml(ByVal strXml As String) As String
    FormatXml = Replace(strXml, "&", "&amp;")
    FormatXml = Replace(FormatXml, "<", "&lt;")
    FormatXml = Replace(FormatXml, ">", "&gt;")
'    FormatXml = Replace(FormatXml, """", "&quot;")
'    FormatXml = Replace(FormatXml, "'", "&quot;")
End Function

Public Function OpenFile(ByVal strFile As String) As Boolean
'打开方案文件
On Error GoTo errHand
    
    Call ReInit
    
    Set mobjXml = Nothing
    Set mobjXml = CreateObject("MSXML2.DOMDocument" & mstrXmlVersion)

    OpenFile = mobjXml.Load(strFile)
    
    If OpenFile Then OpenFile = ParseScheme
    
Exit Function
errHand:
    Set mobjXml = Nothing
    OpenFile = False
End Function

Public Function OpenScheme(ByVal strXml As String) As Boolean
'打开方案
On Error GoTo errHand
    
    Call ReInit
    
    Set mobjXml = Nothing
    Set mobjXml = CreateObject("MSXML2.DOMDocument" & mstrXmlVersion)

    OpenScheme = mobjXml.LoadXML(strXml)
    
    If OpenScheme Then OpenScheme = ParseScheme
Exit Function
errHand:
    Set mobjXml = Nothing
    OpenScheme = False
    Debug.Print "clsSqlScheme.OpenScheme Err:" & Err.Description
End Function

Public Sub SaveAsScheme(ByVal strFile As String)
'保存方案
    Dim objXml As Object
    
    Set objXml = CreateObject("MSXML2.DOMDocument" & mstrXmlVersion)
    
    Call objXml.LoadXML(GetScheme())
    
    Call objXml.Save(strFile)
End Sub


Private Function ParseScheme() As Boolean
'解析方案
    Dim i As Long
    Dim objCurNode As Object
    Dim objCfgNode As Object
    
    Dim objSerachCfg As clsScSerachCfg
    Dim objFilterCfg As clsScFilterCfg
    Dim objShowCfg As clsScShowCfg
'    Dim objRowCfg As clsScRowCfg
    
    ParseScheme = False
    
    Set objCfgNode = mobjXml.selectNodes("scheme").Item(0)
    If objCfgNode Is Nothing Then Exit Function
    
    '读取scheme配置
    mlngSchemeID = GetLngAttribute(objCfgNode, "id")
    mstrSchemeName = GetStrAttribute(objCfgNode, "name")
    mstrDescript = GetStrAttribute(objCfgNode, "descript")
    mdblVer = GetLngAttribute(objCfgNode, "ver")
    
    mblnUseCard = GetBoolAttribute(objCfgNode, "usecard")
    mblnUseGroup = GetBoolAttribute(objCfgNode, "usegroup")
    mblnShowHistory = GetBoolAttribute(objCfgNode, "showhistory")
    mblnUseFuncFollow = GetBoolAttribute(objCfgNode, "usefuncfollow")
    mlngPageRecord = GetLngAttribute(objCfgNode, "pagerecord")
    mlngDateRange = GetLngAttribute(objCfgNode, "daterange")
    
    If mobjXml.firstChild.selectNodes("store").Length > 0 Then
        mstrStore = mobjXml.firstChild.selectNodes("store").Item(0).Text
    End If
    '读取query/detail配置
    mstrQuery = mobjXml.firstChild.selectNodes("query").Item(0).Text
    mstrDetail = mobjXml.firstChild.selectNodes("detail").Item(0).Text
    
    '读取serach配置
    Set objCfgNode = mobjXml.firstChild.selectNodes("serachcfg").Item(0)
    Call ReadSerachCfg(objCfgNode)
    
    '读取filter配置
    Set objCfgNode = mobjXml.firstChild.selectNodes("filtercfg").Item(0)
    Call ReadFilterCfg(objCfgNode)
    
    '读取Show配置
    Set objCfgNode = mobjXml.firstChild.selectNodes("showcfg").Item(0)
    Call ReadShowCfg(objCfgNode)
    
    ParseScheme = True
End Function

Private Sub ReadSerachCfg(objCfgNode As Object)
'读取serach配置
    Dim i As Long
    Dim objCurNode As Object
    Dim objSerachCfg As clsScSerachCfg
    
    '读取serach配置
    If Not objCfgNode Is Nothing Then
        For i = 0 To objCfgNode.ChildNodes.Length - 1
        
            Set objSerachCfg = New clsScSerachCfg
            Set objCurNode = objCfgNode.ChildNodes(i)
            
            objSerachCfg.Name = GetStrAttribute(objCurNode, "name")
            objSerachCfg.InputType = GetLngAttribute(objCurNode, "inputtype")
            objSerachCfg.ControlType = GetLngAttribute(objCurNode, "controltype")
'            objSerachCfg.SourceType = GetLngAttribute(objCurNode, "sourcetype")
            objSerachCfg.LikeWay = GetLngAttribute(objCurNode, "likeway")
            objSerachCfg.Default = GetStrAttribute(objCurNode, "default")
            objSerachCfg.DataFrom = objCurNode.Text
            
            Call AddSerachCfg(objSerachCfg)
        Next i
    End If
End Sub

Private Sub ReadFilterCfg(objCfgNode As Object)
'读取serach配置
    Dim i As Long
    Dim j As Long
    Dim objCurNode As Object
    Dim objSubNode As Object
    Dim objFilterCfg As clsScFilterCfg
    
    '读取filter配置
    If Not objCfgNode Is Nothing Then
        For i = 0 To objCfgNode.ChildNodes.Length - 1
        
            Set objFilterCfg = New clsScFilterCfg
            Set objCurNode = objCfgNode.ChildNodes(i)

            objFilterCfg.Name = GetStrAttribute(objCurNode, "name")
'            objFilterCfg.FilterType = GetLngAttribute(objCurNode, "filtertype")
'            objFilterCfg.SourceType = GetLngAttribute(objCurNode, "sourcetype")
            objFilterCfg.SelectWay = GetLngAttribute(objCurNode, "selectway")
'            objFilterCfg.DataFrom = objCurNode.Text
            
            If objCfgNode.ChildNodes.Length > 0 Then
                For j = 0 To objCurNode.ChildNodes.Length - 1
                    Set objSubNode = objCurNode.ChildNodes(j)
                    
                    If objSubNode.nodetype = NODETYPE_ELEMENT Then
                        objFilterCfg.CustomScript = objSubNode.Text
                    Else
                        objFilterCfg.DataFrom = objSubNode.Text
                    End If
                Next j
            Else
                objFilterCfg.DataFrom = objCfgNode.Text
            End If
            
            Call AddFilterCfg(objFilterCfg)
        Next i
    End If
End Sub

Private Sub ReadShowCfg(objCfgNode As Object)
'读取Col配置
    Dim i As Long
    Dim j As Long
    Dim objCurNode As Object
    Dim objShowCfg As clsScShowCfg
    Dim objRowRelationNode As Object
    Dim objRowRelation As clsScRowRelation
    
    '读取Col配置
    If Not objCfgNode Is Nothing Then
        For i = 0 To objCfgNode.ChildNodes.Length - 1
        
            Set objShowCfg = New clsScShowCfg
            Set objCurNode = objCfgNode.ChildNodes(i)
            
            objShowCfg.Name = GetStrAttribute(objCurNode, "name")
            objShowCfg.Icon = GetStrAttribute(objCurNode, "icon")
            objShowCfg.HiddenCol = GetBoolAttribute(objCurNode, "hiddencol")
            objShowCfg.HiddenTitle = GetBoolAttribute(objCurNode, "hiddentitle")
            objShowCfg.HiddenData = GetBoolAttribute(objCurNode, "hiddendata")
            objShowCfg.IsTotal = GetBoolAttribute(objCurNode, "istotal")
            objShowCfg.UseListLocate = GetBoolAttribute(objCurNode, "uselistlocate")
            objShowCfg.SortContrastCol = GetStrAttribute(objCurNode, "sortcontrastcol")
            
            If objCurNode.ChildNodes.Length > 0 Then
                For j = 0 To objCurNode.ChildNodes.Length - 1
                    Set objRowRelationNode = objCurNode.ChildNodes(j)
                    
                    If objRowRelationNode.nodetype = NODETYPE_ELEMENT Then
                        Set objRowRelation = New clsScRowRelation
                        
                        objRowRelation.TiggerData = GetStrAttribute(objRowRelationNode, "tiggerdata")
                        objRowRelation.Icon = GetStrAttribute(objRowRelationNode, "icon")
                        objRowRelation.IconPerformCol = GetStrAttribute(objRowRelationNode, "iconperformcol")
                        objRowRelation.IsStateIcon = GetBoolAttribute(objRowRelationNode, "isstateicon")
                        objRowRelation.RowFontColor = GetLngAttribute(objRowRelationNode, "rowfc")
                        objRowRelation.RowBackColor = GetLngAttribute(objRowRelationNode, "rowbc")
                        objRowRelation.CellFontColor = GetLngAttribute(objRowRelationNode, "cellfc")
                        objRowRelation.CellBackColor = GetLngAttribute(objRowRelationNode, "cellbc")
                        objRowRelation.ColorPerformCol = GetStrAttribute(objRowRelationNode, "colorperformcol")
                        objRowRelation.FlickerTimeOut = GetLngAttribute(objRowRelationNode, "flickertimeout")
                        objRowRelation.TimeOutReferCol = GetStrAttribute(objRowRelationNode, "timeoutrefercol")
                        
                        Call objShowCfg.AddRowRelation(objRowRelation)
                    Else
                        objShowCfg.DataConvert = objRowRelationNode.Text
                    End If
                Next j
            Else
                objShowCfg.DataConvert = objCurNode.Text
            End If
            
            
            Call AddShowCfg(objShowCfg)
        Next i
    End If
End Sub

Private Function GetStrAttribute(objNode As Object, ByVal strProperty As String, Optional ByVal strDefault As String = "") As String
'获取字符串属性
On Error GoTo errHandle
    Dim objAtrNode As Object
    
    GetStrAttribute = ""
    
    Set objAtrNode = objNode.Attributes.getNamedItem(strProperty)
    If objAtrNode Is Nothing Then
        GetStrAttribute = strDefault
        Exit Function
    End If
    
    GetStrAttribute = objAtrNode.Text
Exit Function
errHandle:
    GetStrAttribute = ""
End Function


Private Function GetLngAttribute(objNode As Object, ByVal strProperty As String, Optional ByVal lngDefault As Long = -1) As Long
'获取数字属性
On Error GoTo errHandle
    Dim objAtrNode As Object
    
    GetLngAttribute = 0
    
    Set objAtrNode = objNode.Attributes.getNamedItem(strProperty)
    If objAtrNode Is Nothing Then
        GetLngAttribute = lngDefault
        Exit Function
    End If
    
    If Len(objAtrNode.Text) <= 0 Then
        GetLngAttribute = lngDefault
    Else
        GetLngAttribute = Val(objAtrNode.Text)
    End If
Exit Function
errHandle:
    GetLngAttribute = 0
End Function


Private Function GetBoolAttribute(objNode As Object, ByVal strProperty As String, Optional ByVal blnDefault As Boolean = False) As Boolean
'获取数字属性
On Error GoTo errHandle
    Dim objAtrNode As Object
    
    GetBoolAttribute = False
    
    Set objAtrNode = objNode.Attributes.getNamedItem(strProperty)
    If objAtrNode Is Nothing Then
        GetBoolAttribute = blnDefault
        Exit Function
    End If
    
    If Len(objAtrNode.Text) <= 0 Then
        GetBoolAttribute = blnDefault
    Else
        GetBoolAttribute = CBool(objAtrNode.Text)
    End If
Exit Function
errHandle:
    GetBoolAttribute = False
End Function


Public Sub AddSerachCfg(objSerachCfg As clsScSerachCfg)
'添加检索项配置
    Dim lngUbound As Long
    
    lngUbound = UBound(mobjSerachCfg) + 1
    ReDim Preserve mobjSerachCfg(lngUbound)
    
    Set mobjSerachCfg(lngUbound) = objSerachCfg
End Sub


Public Sub AddFilterCfg(objFilterCfg As clsScFilterCfg)
'添加过滤项配置
    Dim lngUbound As Long
    
    lngUbound = UBound(mobjFilterCfg) + 1
    ReDim Preserve mobjFilterCfg(lngUbound)
    
    Set mobjFilterCfg(lngUbound) = objFilterCfg
End Sub


Public Sub AddShowCfg(objShowCfg As clsScShowCfg)
'添加过滤项配置
    Dim lngUbound As Long
    
    lngUbound = UBound(mobjShowCfg) + 1
    ReDim Preserve mobjShowCfg(lngUbound)
    
    Set mobjShowCfg(lngUbound) = objShowCfg
End Sub

Private Function GetXMLVersion() As String
'获取xml对应的支持版本
    Dim varXMLVersion As Variant
    Dim strXMLVer As String
    Dim intLoop As Integer
    Dim objXml As Object
    
    On Error GoTo errHand
        
    varXMLVersion = Split(".6.0,.4.0", ",")
    
    On Error Resume Next
        For intLoop = 0 To UBound(varXMLVersion)
            Err = 0
            Set objXml = CreateObject("MSXML2.DOMDocument" & varXMLVersion(intLoop))
            If Err = 0 Then
                strXMLVer = varXMLVersion(intLoop)
                Exit For
            End If
        Next
        
    On Error GoTo errHand
        
    If strXMLVer = "" Then
        MsgBox "创建MSXML2.DOMDocument对象失败"
        Exit Function
    End If
 
    GetXMLVersion = strXMLVer
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
End Function

Public Sub ReInit()
    Set mobjXml = Nothing
    
    ReDim mobjSerachCfg(0)
    ReDim mobjFilterCfg(0)
    ReDim mobjShowCfg(0)
    
    ReDim aryCfgConvertCol(0)
    ReDim aryCfgConvertData(0)
    
    mstrQuery = ""
    mstrDetail = ""
    
    mstrSchemeName = ""
    mdblVer = -1
    mstrDescript = ""
    
    mblnShowHistory = False
    mblnUseCard = False
End Sub

Private Sub Class_Initialize()
    mstrXmlVersion = GetXMLVersion
    Call ReInit
End Sub

Private Sub Class_Terminate()
    Set mobjXml = Nothing
End Sub
