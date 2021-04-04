VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mobjScriptControl As Object
                 
'######################################################################################################################
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function OpenJson(ByVal strJson As String, Optional blnAddItemNode As Boolean = True) As Boolean
    
    On Error GoTo errHand

    If strJson = "" Then Exit Function
    
    
    strJson = Replace(strJson, Chr(13), "chr(13)")
    strJson = Replace(strJson, Chr(10), "chr(10)")
    strJson = Replace(strJson, "\", "\\")
    
    mobjScriptControl.Language = "javascript"
    mobjScriptControl.Reset
    If blnAddItemNode Then
       ' strJson = "{" & """" & "item" & """" & ":" & strJson & "};"
    End If
    Call mobjScriptControl.AddCode("var json=" & strJson)
    OpenJson = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function GetValue(ByVal strElement As String, Optional blnAddJson As Boolean = True) As Variant
    Dim strResult As String
    
    On Error GoTo errHand
    If Not blnAddJson Then
        strResult = mobjScriptControl.Eval(strElement)
    Else
        strResult = mobjScriptControl.Eval("json." & strElement)
    End If
    strResult = Replace(strResult, "chr(13)", Chr(13))
    strResult = Replace(strResult, "chr(10)", Chr(10))
    strResult = Replace(strResult, "\\", "\")
    
    GetValue = strResult
     '------------------------------------------------------------------------------------------------------------------
errHand:
    
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function SetValue(ByVal strElement As String, ByVal strNewValue As String) As String
    
    On Error GoTo errHand
    
    'javascript
    SetValue = mobjScriptControl.Eval("json." & strElement & "='" & strNewValue & "'")
    
    
    Exit Function
errHand:
    
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function GetJson() As String
    
    On Error GoTo errHand
    
    'javascript
    GetJson = mobjScriptControl.Eval("alter(json)")
    
    Exit Function
errHand:
    
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function GetLength(ByVal strElement As String) As Long
    On Error GoTo errHand
    
'    var data = "[{'id':'1','name':'a'},{'id':'2','name':'b'}]";
'    alert(eval(data).length);
    
     GetLength = mobjScriptControl.Eval("json." & strElement).length
     
errHand:
    
End Function


'######################################################################################################################
Private Sub Class_Initialize()
    Set mobjScriptControl = CreateObject("ScriptControl")
End Sub

Private Sub Class_Terminate()
    Set mobjScriptControl = Nothing
End Sub
Public Function JSONParse(ByVal JSONPath As String, ByVal JSONString As String) As Variant
    Dim JSON As Object
    Set JSON = CreateObject("MSScriptControl.ScriptControl")
    JSON.Language = "JScript"
    JSONParse = JSON.Eval("JSON=" & JSONString & ";JSON." & JSONPath & ";")
    Set JSON = Nothing
End Function

