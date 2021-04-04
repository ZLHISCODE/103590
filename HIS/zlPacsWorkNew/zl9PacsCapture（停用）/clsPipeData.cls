VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPipeData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DATA_ITEM_FROM  As String = "<@数据来源>"
Private Const DATA_ITEM_NAME  As String = "<@数据项名>"
Private Const DATA_ITEM_TAG As String = "<@数据标记>"
Private Const DATA_ITEM_RUNDOWN As String = "<@数据摘要>"
Private Const DATA_ITEM_COMMAND As String = "<@命令标识>"

Private mlstData As Object

Private mstrDataFrom As String
Private mstrDataName As String
Private mstrDataTag As String
Private mstrDataRundown As String
Private mstrCommandIdentify As String


Property Get Datas() As Collection
    Set Datas = mlstData
End Property

'DataFrom
Property Get DataFrom() As String
    DataFrom = mstrDataFrom
End Property


Property Let DataFrom(ByVal value As String)
    mstrDataFrom = value
    
    Call SetData(DATA_ITEM_FROM, value)
End Property


'DataName
Property Get DataName() As String
    DataName = mstrDataName
End Property


Property Let DataName(ByVal value As String)
    mstrDataName = value
    
    Call SetData(DATA_ITEM_NAME, value)
End Property


'DataTag
Property Get DataTag() As String
    DataTag = mstrDataTag
End Property

Property Let DataTag(ByVal value As String)
    mstrDataTag = value
    
    Call SetData(DATA_ITEM_TAG, value)
End Property


'Rundown
Property Get DataRundown() As String
    DataRundown = mstrDataRundown
End Property

Property Let DataRundown(ByVal value As String)
    mstrDataRundown = value
    
    Call SetData(DATA_ITEM_RUNDOWN, value)
End Property

'CommandIdentify
Property Get CommandIdentify() As String
    CommandIdentify = mstrCommandIdentify
End Property

Property Let CommandIdentify(ByVal value As String)
    mstrCommandIdentify = value
    
    Call SetData(DATA_ITEM_COMMAND, value)
End Property

Public Sub SetData(ByVal strKey As String, ByVal strData As Variant)
    If mlstData.Exists(strKey) Then
        mlstData(strKey) = strData
    Else
        Call mlstData.Add(strKey, strData)
    End If
End Sub


Public Function GetData(ByVal strKey As String) As Variant
    GetData = ""
    
    If mlstData.Exists(strKey) Then
        GetData = mlstData(strKey)
    End If
End Function


'Private Function GetAdoStructorData() As ADODB.Recordset
'    Dim rsData As New ADODB.Recordset
'    Dim strKey As Variant
'
'    Set GetAdoStructorData = Nothing
'
'    If mlstData.Count <= 0 Then Exit Function
'
'    For Each strKey In mlstData.Keys
'        Call rsData.Fields.Append(strKey, adVariant)
'    Next
'
'    rsData.Open
'    rsData.AddNew
'
'    For Each strKey In mlstData.Keys
'        rsData.Fields(strKey) = mlstData(strKey)
'    Next
'
'    Set GetAdoStructorData = rsData
'End Function

Private Function GetCustomXmlData() As String
    Dim strXml As String
    Dim strKey As Variant
    Dim strFormatValue As String
    Dim strFormatKey As String
    
    strXml = ""
    
    If mlstData.Count <= 0 Then Exit Function
    
    For Each strKey In mlstData.Keys
        strFormatValue = mlstData(strKey)
'        strFormatValue = Replace(strFormatValue, "<", "&lt;")
'        strFormatValue = Replace(strFormatValue, ">", "&gt;")
         
        strFormatKey = Replace(strKey, "<", "_x003C_")
        strFormatKey = Replace(strFormatKey, "@", "_x0040_")
        
        strFormatKey = Replace(strFormatKey, ">", "_x003E_")
        
        strXml = strXml + "<" + strFormatKey + ">" + strFormatValue + "</" + strFormatKey + ">"
    Next
    
    GetCustomXmlData = "<xml>" + strXml + "</xml>"
End Function


'保存为xml格式数据
Public Function SaveToXml() As String
'    Dim rsData As ADODB.Recordset
'    Dim stream As New ADODB.stream
'
'    SaveToXml = ""
'    Set rsData = GetAdoStructorData()
'
'    stream.Type = adTypeText
'    stream.Mode = adModeReadWrite
'    stream.Open
'
'    Call rsData.Save(stream, adPersistXML)
'    stream.Position = 0
'
'    SaveToXml = stream.ReadText()  '性能很低
    
    SaveToXml = GetCustomXmlData()
End Function



'保存数据为xml文件
Public Sub SaveToXmlFile(ByVal strFile As String)
    Dim strXml As String
    strXml = SaveToXml()
    
    Call StrToFile(strFile, strXml)
End Sub


'Public Sub LoadXmlFile(ByVal strFile As String)
'    Dim stream As New ADODB.stream
'    Dim rsData As New ADODB.Recordset
'    Dim strField As Variant
'
'    stream.Type = adTypeBinary
'    stream.Mode = adModeReadWrite
'
'    stream.Open
'    Call stream.LoadFromFile(strFile)
'    stream.Position = 0
'
'    rsData.Open stream
'
'    For Each strField In rsData.Fields
'        Call SetData(strField, rsData(strField).value)
'
'        Call ParseStructure(strField, rsData(strField).value)
'    Next
'
'End Sub

'Private Sub ParseStructure(ByVal structureName As String, ByVal value As String)
'    If UCase(structureName) = DATA_ITEM_FROM Then
'        mstrDataFrom = value
'    End If

'    If UCase(structureName) = DATA_ITEM_NAME Then
'        mstrDataName = value
'    End If
'
'    If UCase(structureName) = DATA_ITEM_TAG Then
'        mstrDataTag = value
'    End If
'
'    If UCase(structureName) = DATA_ITEM_RUNDOWN Then
'        mstrDataRundown = value
'    End If

'    If UCase(structureName) = DATA_ITEM_COMMAND Then
'        mstrCommandIndetify = value
'    End If
'End Sub

Private Sub StrToFile(ByVal strFile As String, ByVal strData As String)
    Open strFile For Output As #1
     Print #1, strData
     Close #1
End Sub

Private Sub Class_Initialize()
    Set mlstData = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set mlstData = Nothing
End Sub
