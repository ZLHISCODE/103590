Attribute VB_Name = "mdlUserProcedure"
Option Explicit

'1-�䶯����;2-�հ׹���;3-�û�����
Public Enum ProcType
    �䶯���� = 1
    �հ׹��� = 2
    �û����� = 3
End Enum

Public Enum ProcState
    ������ = 1
    ������ = 2
    �ѵ��� = 3
End Enum

'1-�ϴ��Զ�����;2-�ϴα�׼����;3-�����Զ�����;4-���α�׼����
Public Enum ProcTextType
    �ϴ��Զ����� = 1
    �ϴα�׼���� = 2
    �����Զ����� = 3
    ���α�׼���� = 4
End Enum

Public Enum Color
    ��ɫ = &H80000005
    ��ɫ = &HFF&
    ��ɫ = &HFF0000
    ��ɫ = 0
    �ǽ��� = &HFFEBD7
    ���� = &HFFCC99
    ǳ��ɫ = &HE0E4E7
    ���ɫ = &H8000000C
    ��ɫ = &H8000000F
    ǳ��ɫ = &H80000018
    

    
    ����ģ��ɫ = &HC00000

    
    Ĭ��ǰ��ɫ = &H80000008
    ��ɫ = &HF5F5F5
    ����ɫ = 0
    ͣ��ɫ = 255
End Enum

Public gstrSplite As String

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
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
    
    strSQL = "Select 1 From zlProcedure Where ����=[1] And ����=2"
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
'    '���As/Isǰ�벿��
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
'    '��ý�������
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
        
        'ȡ��--ע��
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





