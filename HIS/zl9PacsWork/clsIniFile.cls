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
'˵������дINI�����ļ�
'�������ڣ�2010-01-12
'�����ˣ�TJH
'
'����������
'
'*********************************************************************************************************
'---------------------------------------------------------------------------------------------------------


Option Explicit
    
Private mstrFileName As String
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'����INI�ļ�����
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIniFile(ByVal strFileName As String)
  mstrFileName = strFileName
End Sub

    
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'FileName��Ini�ļ�
'PathName��С����
'KeyName��ֵ��
'WriteValue��ֵ
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function WriteValue(strSectionName As String, strKeyName As String, strWriteValue As String) As Boolean
  Dim lngWriteState As Long
            
  lngWriteState = WritePrivateProfileString(strSectionName, strKeyName, strWriteValue, mstrFileName)
            
  WriteValue = (lngWriteState <> 0)
            
End Function
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'FileName��Ini�ļ�
'PathName��С����
'KeyName��ֵ��
'BackValue������ֵ
'Default��Ĭ���ַ�
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
    ReadValue = Trim(Left$(strTemp, strTempNum))
  End If
            
End Function
