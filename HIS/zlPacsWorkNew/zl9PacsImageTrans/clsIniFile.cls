VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------------------------
'*********************************************************************************************************
'
'说明：读写INI配置文件
'创建日期：2010-01-12
'创建人：TJH
'
'其他描述：
'
'*********************************************************************************************************
'---------------------------------------------------------------------------------------------------------


Option Explicit
    
Private mstrFileName As String
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'设置INI文件名称
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIniFile(ByVal strFileName As String)
  mstrFileName = strFileName
End Sub

    
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'FileName：Ini文件
'PathName：小节名
'KeyName：值名
'WriteValue：值
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function WriteValue(strSectionName As String, strKeyName As String, strWriteValue As String) As Boolean
  Dim lngWriteState As Long
            
  lngWriteState = WritePrivateProfileString(strSectionName, strKeyName, strWriteValue, mstrFileName)
            
  WriteValue = (lngWriteState <> 0)
            
End Function
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'FileName：Ini文件
'PathName：小节名
'KeyName：值名
'BackValue：返回值
'Default：默认字符
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ReadValue(strSectionName As String, strKeyName As String, _
                          Optional strDefault As String = "") As String
  Dim lngReadState As Long
  Dim strTempNum As String
  Dim strTemp As String
            
  strTemp = String$(255, Chr$(0))
  strTempNum = 255
  
  ReadValue = strDefault
            
  lngReadState = GetPrivateProfileString(strSectionName, strKeyName, strDefault, strTemp, strTempNum, mstrFileName)
                        
  If lngReadState <> 0 Then
    ReadValue = Replace((Left$(strTemp, strTempNum)), Chr$(0), "")
  End If
            
End Function

