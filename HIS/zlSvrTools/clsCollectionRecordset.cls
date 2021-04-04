VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCollectionRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'变量定义
Private mstrSplit As String
Private mobjCollection As New Collection
Private mstrFieldValue As String
Private mvarFieldValue As Variant
Private mobjCollectionField As New Collection
Private mlngCount As Long

Private mintDataFieldCount As Integer
Private mintKeyFieldCount As Integer
Private mlngSplitLen As Long

'######################################################################################################################

Public Sub DataFields(ByVal strDataField As Parameter)
    
End Sub

Public Sub Add(ByVal strKey As String, ByVal strItem As String)
                
    mobjCollection.Add strItem, "K" & strKey

    mlngCount = mobjCollection.Count

End Sub

Public Function LoadData(ByVal rsData As ADODB.Recordset, ByVal varDataField As Variant, ByVal varKeyField As Variant) As Boolean
    
    Dim strKey As String
    Dim strItem As String
    Dim intLoop As Integer
    Dim intCount1 As Integer
    Dim intCount2 As Integer
    
    mstrFieldValue = ""
        
    mlngCount = 0
    
    mintKeyFieldCount = UBound(varKeyField)
    If mintKeyFieldCount = -1 Then Exit Function
    
    mintDataFieldCount = UBound(varDataField)
    mlngSplitLen = Len(mstrSplit)
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst

        For intLoop = 0 To mintDataFieldCount
            mobjCollectionField.Add CStr(intLoop), "K" & UCase(varDataField(intLoop))
        Next
        
        
        Do While Not rsData.EOF
            
            strItem = ""
            If mintDataFieldCount = -1 Then
                strItem = mstrSplit & "1"
            Else
                For intLoop = 0 To mintDataFieldCount
                    strItem = strItem & mstrSplit & rsData(varDataField(intLoop)).value
                Next
            End If
            
            strKey = ""
            For intLoop = 0 To mintKeyFieldCount
                strKey = strKey & mstrSplit & rsData(varKeyField(intLoop)).value
            Next
            
            If strItem <> "" And strKey <> "" Then
                strItem = Mid(strItem, mlngSplitLen + 1)
                strKey = Mid(strKey, mlngSplitLen + 1)
                mobjCollection.Add strItem, "K" & strKey
            End If
            rsData.MoveNext
        Loop
        
        mlngCount = mobjCollection.Count
    End If
End Function

Public Property Get Count() As Long
    Count = mlngCount
End Property

Public Function Item(ByVal varKey As Variant) As Boolean
    Dim strKey As String
    
    On Error Resume Next
    
    strKey = Join(varKey, mstrSplit)
    
    mstrFieldValue = ""
    mstrFieldValue = mobjCollection.Item("K" & strKey)
    If mstrFieldValue <> "" Then
        mvarFieldValue = Split(mstrFieldValue, mstrSplit)
    End If
    If err.Number <> 0 Then err.Clear
    Item = (mstrFieldValue <> "")
    
End Function

Public Function ItemValue(ByVal strField As String) As String
    Dim intIndex As Integer
    
    On Error Resume Next
    
    intIndex = -1
    intIndex = Val(mobjCollectionField("K" & UCase(strField)))
    
    On Error GoTo 0
    
    If intIndex >= 0 Then
        ItemValue = mvarFieldValue(intIndex)
    End If
    
End Function

Private Sub Class_Initialize()
    mstrSplit = "IAMASPLITSTRING"
End Sub
