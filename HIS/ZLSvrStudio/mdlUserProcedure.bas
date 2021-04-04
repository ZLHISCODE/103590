Attribute VB_Name = "mdlUserProcedure"
Option Explicit

'1-变动过程;2-空白过程;3-用户过程
Public Enum ProcType
    变动过程 = 1
    空白过程 = 2
    用户过程 = 3
End Enum

Public Enum ProcState
    待调整 = 1
    调整中 = 2
    已调整 = 3
End Enum

'1-上次自定过程;2-上次标准过程;3-本次自定过程;4-本次标准过程
Public Enum ProcTextType
    上次自定过程 = 1
    上次标准过程 = 2
    本次自定过程 = 3
    本次标准过程 = 4
End Enum

Public Enum Color
    白色 = &H80000005
    红色 = &HFF&
    兰色 = &HFF0000
    黑色 = 0
    非焦点 = &HFFEBD7
    焦点 = &HFFCC99
    浅灰色 = &HE0E4E7
    深灰色 = &H8000000C
    灰色 = &H8000000F
    浅黄色 = &H80000018
    

    
    公共模块色 = &HC00000

    
    默认前景色 = &H80000008
    锁色 = &HF5F5F5
    启用色 = 0
    停用色 = 255
End Enum

Public gstrSplite As String

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Sub ShowAbout(Optional frmParent As Object)
    Dim frmShow As New frmAbout
    If frmParent Is Nothing Then
        frmShow.Show 1
    Else
        Load frmShow
        err.Clear
        On Error Resume Next
        frmShow.Show 1, frmParent
        If err.Number <> 0 Then
            err.Clear
            frmShow.Show 1
        End If
    End If
End Sub

Public Function IsSpaceProcedure(ByVal strOwner As String, ByVal strProcName As String) As Boolean
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "Select 1 From zlProcedure Where 名称=[1] And 类型=2"
    Set rsData = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", UCase(strProcName))
    IsSpaceProcedure = (rsData.BOF = False)
    
End Function

'Public Function GetSpaceProcedure(ByVal strProc As String) As String
'    Dim lngCount As Long
'    Dim blnTitleFlag As Boolean
'    Dim strEnd As String
'    Dim strSQL As String
'    Dim lngInstr As Long
'    Dim strArr() As String
'
'
'
'    '获得As/Is前半部分
'    Select Case True
'    Case InStr(UCase(strProc), " IS" & vbCrLf) > 0
'        strSQL = Mid(strProc, 1, InStr(UCase(strProc), " IS" & vbCrLf) + Len(" IS"))
'    Case InStr(UCase(strProc), " IS" & Chr(10)) > 0
'        strSQL = Mid(strProc, 1, InStr(UCase(strProc), " IS" & Chr(10)) + Len(" IS"))
'    Case InStr(UCase(strProc), " IS" & Chr(13)) > 0
'        strSQL = Mid(strProc, 1, InStr(UCase(strProc), " IS" & Chr(13)) + Len(" IS"))
'    Case InStr(UCase(strProc), " AS" & vbCrLf) > 0
'        strSQL = Mid(strProc, 1, InStr(UCase(strProc), " AS" & vbCrLf) + Len(" AS"))
'    Case InStr(UCase(strProc), " AS" & Chr(10)) > 0
'        strSQL = Mid(strProc, 1, InStr(UCase(strProc), " AS" & Chr(10)) + Len(" AS"))
'    Case InStr(UCase(strProc), " AS" & Chr(13)) > 0
'        strSQL = Mid(strProc, 1, InStr(UCase(strProc), " AS" & Chr(13)) + Len(" AS"))
'    End Select
'
'
'    strSQL = strSQL & "  Begin"
'    '获得结束部分
'    strArr = Split(strProc, vbCrLf)
'    For lngCount = UBound(strArr) To 0 Step -1
'        If blnTitleFlag = False Then
'            If InStr(UCase(strArr(lngCount)), "END") > 0 Then
'                strEnd = strArr(lngCount)
'                blnTitleFlag = True
'            End If
'        End If
'        If InStr(UCase(strArr(lngCount)), "RETURN") > 0 Then
'            strSQL = strSQL & vbCrLf & "  Return '';"
'            Exit For
'        End If
'    Next
'    strSQL = strSQL & vbCrLf & strEnd
'    GetSpaceProcedure = strSQL
'End Function

Private Function TrimChar(ByVal strText As String) As String
    strText = Replace(strText, Chr(10), "")
    strText = Replace(strText, Chr(13), "")
    TrimChar = strText
End Function

Public Function GetBlankProcedure(ByVal strProc As String) As String
    Dim lngCount As Long
'    Dim blnTitleFlag As Boolean
'    Dim strEnd As String
    Dim strSQL As String
'    Dim lngInstr As Long
    Dim strArr() As String
    
    Dim strLine As String
    Dim lngPostion As Long
    
    Dim strReturnType As String
    
    strArr = Split(strProc, vbCrLf)
    strSQL = ""
    strReturnType = ""
    
    For lngCount = 0 To UBound(strArr)
        
        strLine = UCase(Trim(strArr(lngCount)))
        strLine = TrimChar(strLine)
        
        '取掉--注释
        lngPostion = InStr(strLine, "--")
        If lngPostion > 0 Then strLine = Mid(strLine, 1, lngPostion - 1)
        
        lngPostion = InStr(strLine, "RETURN ")
        If lngPostion > 0 Then
            
            If InStr(strLine, " NUMBER") > 0 Then
                strReturnType = "NUMBER"
            ElseIf InStr(strLine, " VARCHAR") > 0 Then
                strReturnType = "VARCHAR"
            ElseIf InStr(strLine, " DATE") > 0 Then
                strReturnType = "DATE"
            End If
        End If
        
        Select Case strLine
        Case "AS", "IS"
            strSQL = strSQL & strArr(lngCount) & vbCrLf
            Exit For
        Case Else
            If Right(strLine, 3) = " AS" Then
                strSQL = strSQL & strArr(lngCount) & vbCrLf
                Exit For
            ElseIf Right(strLine, 3) = " IS" Then
                strSQL = strSQL & strArr(lngCount) & vbCrLf
                Exit For
            Else
                strSQL = strSQL & strArr(lngCount) & vbCrLf
            End If
        
        End Select
        

    Next
    
    strSQL = strSQL & "Begin" & vbCrLf
    strSQL = strSQL & " " & vbCrLf
    
    Select Case strReturnType
    Case "NUMBER"
        strSQL = strSQL & vbTab & "return 0;" & vbCrLf
    Case "VARCHAR"
        strSQL = strSQL & vbTab & "return '';" & vbCrLf
    Case "DATE"
        strSQL = strSQL & vbTab & "return sysdate;" & vbCrLf
    End Select
    
       
    strSQL = strSQL & "End;"
    
    GetBlankProcedure = strSQL
End Function





