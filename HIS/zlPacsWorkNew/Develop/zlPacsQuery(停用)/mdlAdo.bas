Attribute VB_Name = "mdlAdo"
Option Explicit

Public gcnOracle As ADODB.Connection


Public Function ReadSchemeXml(ByVal lngSchemeId As Long, ByRef strSchemeDes As String) As String
'读取方案的xml配置结构
    Dim lngBaseCount As Long
    Dim strSchemeXml As String
    Dim lngSchemeIndex As Long
    
    ReadSchemeXml = ""
    lngBaseCount = 0
    Do While True
        strSchemeXml = strSchemeXml & ReadSchemeSectionStr(lngSchemeId, strSchemeDes, lngBaseCount) '0-20000字节
        If Trim(strSchemeXml) = "" Then Exit Do
        
        '判断方案是否读取完整
        lngSchemeIndex = InStrRev(strSchemeXml, "</scheme>")
        If lngSchemeIndex <= 0 Then
            lngBaseCount = lngBaseCount + 1
        Else
            If Trim$(Replace(Mid$(strSchemeXml, lngSchemeIndex + 9), vbCrLf, "")) <> "" Then
                lngBaseCount = lngBaseCount + 1
            Else
                Exit Do
            End If
        End If
        
        If lngBaseCount > 5 Then Exit Do
    Loop
    
    ReadSchemeXml = strSchemeXml
        
End Function


Private Function ReadSchemeSectionStr(ByVal lngSchemeId As Long, ByRef strSchemeDes As String, _
    Optional ByVal lngStartBase As Long = 0) As String
'从数据库读取方案的xml字符
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strSchemeXml As String
    Dim lngCurBase As Long
    
    lngCurBase = lngStartBase * 10
    
    ReadSchemeSectionStr = ""
    
    strSql = "Select A.ID, A.方案说明, DBMS_LOB.Substr(方案内容,2000," & 1 + lngCurBase * 2000 & ") As s1," & _
                            " DBMS_LOB.Substr(方案内容,2000," & 1 + (lngCurBase + 1) * 2000 & ") As s2, " & _
                            " DBMS_LOB.Substr(方案内容,2000," & 1 + (lngCurBase + 2) * 2000 & ") As s3, " & _
                            " DBMS_LOB.Substr(方案内容,2000," & 1 + (lngCurBase + 3) * 2000 & ") As s4, " & _
                            " DBMS_LOB.Substr(方案内容,2000," & 1 + (lngCurBase + 4) * 2000 & ") As s5, " & _
                            " DBMS_LOB.Substr(方案内容,2000," & 1 + (lngCurBase + 5) * 2000 & ") As s6, " & _
                            " DBMS_LOB.Substr(方案内容,2000," & 1 + (lngCurBase + 6) * 2000 & ") As s7, " & _
                            " DBMS_LOB.Substr(方案内容,2000," & 1 + (lngCurBase + 7) * 2000 & ") As s8, " & _
                            " DBMS_LOB.Substr(方案内容,2000," & 1 + (lngCurBase + 8) * 2000 & ") As s9, " & _
                            " DBMS_LOB.Substr(方案内容,2000," & 1 + (lngCurBase + 9) * 2000 & ") As s10 " & _
            " From 影像查询方案 A Where id=[1]"
    
    Set rsData = ExecuteSql(strSql, "查询方案", lngSchemeId)
    
    If rsData Is Nothing Then Exit Function
    If rsData.RecordCount <= 0 Then Exit Function
    
    strSchemeDes = NVL(rsData!方案说明)
    
    strSchemeXml = NVL(rsData!s1) & NVL(rsData!s2) & _
                   NVL(rsData!s3) & NVL(rsData!s4) & _
                   NVL(rsData!s5) & NVL(rsData!s6) & _
                   NVL(rsData!s7) & NVL(rsData!s8) & _
                   NVL(rsData!s9) & NVL(rsData!s10)
                   
    ReadSchemeSectionStr = strSchemeXml
End Function

Public Function HasField(ByRef rsData As ADODB.Recordset, ByVal strFieldName As String) As Boolean
On Error GoTo errHandle
    HasField = False
    
    If rsData.Fields(strFieldName).Type <> -1 Then
        HasField = True
    End If
Exit Function
errHandle:
    HasField = False
End Function

Public Function ExecuteSql(ByVal strSql As String, ByVal strTitle As String, _
    ParamArray arrInput() As Variant) As ADODB.Recordset
'执行sql查询
    Dim varPars() As Variant
    
    varPars = arrInput
    
    Set ExecuteSql = ExecuteCore(strSql, strTitle, varPars, False, False)
End Function


Public Sub ExecuteCmd(ByVal strSql As String, ByVal strTitle As String)
'执行存储过程
    Dim varPars() As Variant
    
    Call ExecuteCore(strSql, strTitle, varPars, True)
End Sub


Public Function ExecuteCore(ByVal strSql As String, _
    ByVal strTitle As String, arrInput() As Variant, _
    Optional ByVal blnIsStore = False, Optional ByVal blnStart0 As Boolean = True) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
    Dim cmdData As New ADODB.Command
    Dim rsData As New ADODB.Recordset
    Dim arrPar() As Long
    Dim intMax As Integer
    Dim i As Integer
    Dim lngParCount As Long
    Dim varValue As Variant

    '分析自定的[x]参数
    Call GetParNos(strSql, arrPar)

    '替换为"?"参数
    lngParCount = UBound(arrPar)
    
    For i = 1 To lngParCount
        strSql = Replace(strSql, "[" & arrPar(i) & "]", "?")
        varValue = arrInput(IIf(blnStart0, arrPar(i), arrPar(i) - 1))
        
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        End Select
    Next i

    Call ConfigCommandConnection(cmdData, strTitle)
    If cmdData.ActiveConnection Is Nothing Then Exit Function
    
    cmdData.CommandText = strSql
    If blnIsStore Then
        Call cmdData.Execute
        
        Set ExecuteCore = Nothing
    Else
        rsData.CursorLocation = adUseClient
        rsData.CursorType = adOpenDynamic
        rsData.LockType = adLockOptimistic
        
        rsData.Open cmdData
        
        Set rsData.ActiveConnection = Nothing
        Set ExecuteCore = rsData
    End If
    
'    Set ExecuteCore = cmdData.Execute
'    Set ExecuteCore.ActiveConnection = Nothing
 
End Function

Private Sub ConfigCommandConnection(ByRef cmdData As ADODB.Command, Optional ByVal strTitle As String = "")
On Error Resume Next
    Do While True
        Set cmdData.ActiveConnection = gcnOracle '这句比较慢(这句执行1000次约0.5x秒)
        If Err.Number <> 0 Then
            If MsgBox("数据服务连接异常，是否重试？" & vbCrLf & _
                        "错误原因：" & Err.Description & vbCrLf & _
                        "错误源：" & Err.Source, vbYesNo, IIf(Len(strTitle) > 0, strTitle, "警告")) = vbYes Then
                gcnOracle.Close
                gcnOracle.Open
                
                Err.Clear
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    Loop
End Sub

Public Function zlIconResRead(ByVal strResName As String) As String
    Const conChunkSize As Integer = 10240
    
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim rsLob As New ADODB.Recordset
    Dim strSql As String
    Dim strCurFile As String
    
    Err = 0: On Error GoTo errHand
    
    lngFileNum = FreeFile
    
    strCurFile = gstrCachePath & "\" & strResName & ".bmp"
    strCurFile = Replace(strCurFile, "\\", "\")
    
    If Len(Dir(strCurFile)) <> 0 Then
        zlIconResRead = strCurFile
        Exit Function
    End If
 
 
    Open strCurFile For Binary As lngFileNum
    
    strSql = "Select Zl_影像查询_读取图标('" & strResName & "'," & "[1]) as 图标 From Dual"
    
    lngCount = 0
    Do
        Set rsLob = ExecuteSql(strSql, "zlBlobRead", lngCount)
        
        If rsLob.EOF Then Exit Do
        
        If IsNull(rsLob.Fields(0).value) Then Exit Do
        strText = rsLob.Fields(0).value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    
    Close lngFileNum
    If lngCount = 0 Then Kill strCurFile: strCurFile = ""
    
    zlIconResRead = strCurFile
    
    Exit Function

errHand:
    Close lngFileNum
    Kill strCurFile: zlIconResRead = ""
End Function

Public Function RecordSetToXml(ByRef rsData As ADODB.Recordset) As String
'recordset转xml字符
    Dim strDataSource As String
    Dim strReadText As String
    Dim strSchema As String
    Dim lngStartPos As Long
    Dim strNewPro As String
    Dim strXml As String
    Dim strXML_RSDATA As String
    Dim strXML_SCHEMA As String
    Dim i As Long
    Dim objField As ADODB.Field
    Dim strNames As String
    Dim strSplitName As String
    Dim strBufer(100) As String
    Dim lngBufCount As Long
    
    Dim adoSourceStream As ADODB.Stream
    
    Set adoSourceStream = New ADODB.Stream
    adoSourceStream.Type = adTypeText
    adoSourceStream.Mode = adModeRead
    
    Call rsData.Save(adoSourceStream, adPersistXML)
    
    adoSourceStream.Position = 0
    
    lngBufCount = 0
    '如果一次性读取所有数据，速度会变很慢
    strReadText = adoSourceStream.ReadText(200000)
    While Trim(strReadText) <> ""
        strBufer(lngBufCount) = strBufer(lngBufCount) & strReadText
        If Len(strBufer(lngBufCount)) >= 2000000 Then
            lngBufCount = lngBufCount + 1
        End If
        
        strReadText = adoSourceStream.ReadText(200000)
    Wend
    
    strDataSource = ""

    For i = 0 To lngBufCount
        strDataSource = strDataSource & strBufer(i)
    Next i
    
    lngStartPos = InStr(strDataSource, "<rs:data>")
    strXML_RSDATA = Mid(strDataSource, lngStartPos, InStr(strDataSource, "</rs:data>") - lngStartPos + 11)
    
    '将日期格式为“2012-03-04T12:13:14”替换为“2012-03-04 12:13:14”
'    strXML_RSDATA = RegReplace(strXML_RSDATA, "(?!\b-\d{1,2})T(?=\d{1,2}:)", " ")
    
    strNewPro = ""
    '增加绑定显示的数据列
    For i = 0 To rsData.Fields.Count - 1
        Set objField = rsData.Fields(i)
        strSplitName = "," & objField.Name & ","
        
        If InStr(strNames, strSplitName) <= 0 Then
            strNames = strNames & strSplitName
        
            If strNewPro <> "" Then strNewPro = strNewPro & vbCrLf
            If objField.Type = adDate Or objField.Type = adDBTimeStamp Or objField.Type = adDBDate Or objField.Type = adDBTime Then
                strNewPro = strNewPro & "<s:AttributeType name='" & objField.Name & "' rs:number='" & i + 1 & "' rs:nullable='true' rs:writeunknown='true'>" & _
                        "<s:datatype dt:type='dateTime' rs:dbtype='timestamp' rs:scale='0' rs:precision='3' rs:fixedlength='true'/>" & _
                        "</s:AttributeType>"
            Else
                strNewPro = strNewPro & "<s:AttributeType name='" & objField.Name & "' rs:number='" & i + 1 & "' rs:nullable='true' rs:writeunknown='true'>" & _
                        "<s:datatype dt:type='string' rs:dbtype='str' rs:scale='0' rs:precision='3' rs:fixedlength='true'/>" & _
                        "</s:AttributeType>"
            End If
        End If
    Next i
    
    strXML_SCHEMA = "<s:Schema id='RowsetSchema'>" & vbCrLf & _
                    "    <s:ElementType name='row' content='eltOnly' rs:CommandTimeout='30'" & vbCrLf & _
                    "    rs:updatable='true' rs:ReshapeName='DSRowset1'>" & vbCrLf & _
                    strNewPro & vbCrLf & _
                    "    <s:extends type='rs:rowbase'/>" & vbCrLf & _
                    "</s:ElementType>" & vbCrLf & _
                    "</s:Schema>"
                    
    strXml = "<xml xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882'" & vbCrLf & _
            "xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'" & vbCrLf & _
            "xmlns:rs='urn:schemas-microsoft-com:rowset'" & vbCrLf & _
            "xmlns:z='#RowsetSchema'>" & vbCrLf & _
            strXML_SCHEMA & vbCrLf & _
            strXML_RSDATA & vbCrLf & _
            "</xml>"
            
    RecordSetToXml = strXml
End Function

Public Function XmlToRecordSet(ByRef strData As String) As ADODB.Recordset
    Dim adoNewStream As ADODB.Stream
    Dim rsData As New ADODB.Recordset
    
    Set adoNewStream = New ADODB.Stream
    adoNewStream.Type = adTypeText
    adoNewStream.Mode = adModeWrite
    
    '读取修改后的流数据
    adoNewStream.Open
    adoNewStream.WriteText strData
    adoNewStream.Position = 0
    
    rsData.CursorLocation = adUseClient
    rsData.CursorType = adOpenDynamic
    rsData.LockType = adLockOptimistic
    
    rsData.Open adoNewStream
    
    Set XmlToRecordSet = rsData
End Function

Public Function CopyRecordSet(ByRef rsData As ADODB.Recordset) As ADODB.Recordset
'复制新的数据集使其可以修改
    Dim rsNew As ADODB.Recordset
    Dim strXml As String
    
    Set CopyRecordSet = Nothing
    
    If rsData Is Nothing Then Exit Function
    
    strXml = RecordSetToXml(rsData)
    Set rsNew = XmlToRecordSet(strXml)
    
    Set CopyRecordSet = rsNew
End Function

Public Function zlBlobSave(ByVal KeyWord As String, ByVal strFile As String) As Boolean
'保存图标
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim strSql As String
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)

    Err = 0: On Error GoTo errHand

    conChunkSize = 2000
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If

        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next

        strText = Join(aryHex, "")
        strSql = "Zl_影像查询_保存图标('" & KeyWord & "','" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"

        Call ExecuteCmd(strSql, "zlBlobSave")
    Next
    Close lngFileNum
    zlBlobSave = True

    Exit Function

errHand:
    Close lngFileNum
    zlBlobSave = False
End Function

Public Function zlBlobRead(ByVal KeyWord As String, Optional strFile As String) As String
'读取图标
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim rsLob As New ADODB.Recordset
    Dim strSql As String
    
    Err = 0: On Error GoTo errHand
    
    lngFileNum = FreeFile
    If strFile = "" Then
        lngCount = 0
        Do While True
            strFile = App.Path & "\zlBlobFile" & CStr(lngCount) & ".ico"
            If Len(Dir(strFile)) = 0 Then
                Exit Do
            End If
            lngCount = lngCount + 1
        Loop
    End If
    Open strFile For Binary As lngFileNum
    
    strSql = "Select Zl_影像查询_读取图标('" & KeyWord & "'," & "[1]) as 图标 From Dual"
    lngCount = 0
    Do
        Set rsLob = ExecuteSql(strSql, "zlBlobRead", lngCount)
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).value) Then Exit Do
        strText = rsLob.Fields(0).value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    Close lngFileNum
    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile
    Exit Function

errHand:
    Close lngFileNum
    Kill strFile: zlBlobRead = ""
End Function
