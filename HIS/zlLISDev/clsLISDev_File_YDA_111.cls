VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_YDA_111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:YDA-111 血气仪
'厂商:
'作者:李建飞
'用户:贵阳公司-王伟(思蓝县人民医院)
Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
'标准函数：从文件获取检验结果
'   strFile：包含路径的文件名
'   lngSampleNO：标本号。为-1表示选取所有时间范围内的标本
'   dtStart：开始时间。如果指定了标本号，则只提取该参数当天的这个标本（dtEnd无效）
'   dtEnd：结束时间。只有当选取多个标本（lngSampleNO=-1）时，该参数有效。如果不指定该参数值，则时间范围为>=dtStart。
'   返回的检验结果数组(各解析程序必须按以下标准组织结果)
'   元素之间以|分隔
'   第0个元素：检验时间
'   第1个元素：样本序号
'   第2个元素：检验人
'   第3个元素：标本
'   第4个元素：是否质控品
'   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
'       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    Dim i As Integer
    Dim aResults() As String
    Dim frmPic As New frmGraph
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsAnti As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim str结果 As String, str标本号 As String
    Dim lngResultCount As Long
    Dim str参数 As String, arr参数() As String
    Dim str日期 As String, iField As Integer
    Dim aPic() As String
    Dim strPic As String
    Dim strPicPath As String               '图像访问路径
    Dim arrHigh(2) As Double, arrVAL(2) As Double, arrLow(2) As Double
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    '配置文件用的是仪器商的配置文件，在技师站的参数选择中指定文件时，指定此配置文件，配置文件内容中要用到的内容如下：
    '[BDE_MSQL_XLB]
    'Database Name = th_xlb2035
    'SERVER Name = PC081128
    'USER Name = sa
    'PASSWORD=
    '--------------------------------------------------------------------------------------------------------------------
'
    strPicPath = ReadIni("YDA-111", "图片路径", App.Path & "\YDA-111.ini")
    If strPicPath = "" Then strPicPath = "E:": Call WriteIni("YDA-111", "图片路径", strPicPath, App.Path & "\YDA-111.ini")
    '读配置文件
    Call ReadConf(strFile, "SERVER NAME,DATABASE NAME,USER NAME,PASSWORD", str参数)
    arr参数 = Split(str参数, ",")
    If UBound(arr参数) <> 3 Then Exit Function
    With cnAccess
        .ConnectionString = _
        "Provider=SQLOLEDB.1;Initial Catalog=" & arr参数(1) & ";Password=" & arr参数(3) & _
        ";Persist Security Info=True;User ID=" & arr参数(2) & ";Data Source=" & arr参数(0)
        .Open
     
        If .State = adStateClosed Then Exit Function
        strSQL = "SELECT  * from persons" & _
                " where check_date>='" & Format(dtStart, "yyyy-MM-dd") & "'" & _
                " And check_date<='" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
                IIf(lngSampleNO = -1, "", " And person_ID=" & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If

        ReDim aResults(0)
        lngResultCount = 0

        Do Until rsTmp.EOF
            If Val("" & rsTmp("person_ID")) > 0 Then '新标本
                str标本号 = "" & rsTmp("person_ID")
                str日期 = Format(rsTmp("check_Date"), "YYYY-MM-DD")        '检验时间
                str结果 = ""
                
                For i = 0 To 2
                    arrHigh(i) = rsTmp("id" & i + 2 & "_high")
                    arrVAL(i) = rsTmp("id" & i + 2 & "_val")
                    arrLow(i) = rsTmp("id" & i + 2 & "_low")
                Next
                Call GetResult(arrHigh()): Call GetResult(arrVAL()): Call GetResult(arrLow())
                
                For iField = 12 To 41
                    If "" & rsTmp.Fields(iField).Value <> "" Then
                        str结果 = str结果 & "|" & rsTmp.Fields(iField).Name & "|" & rsTmp.Fields(iField + 30).Value
                    End If
                Next
            
                If str标本号 <> "" And str结果 <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = str日期 & "|" & Val(str标本号) & "| |血清|0" & str结果
                    strPic = "粘度特征曲线;110;1;" & frmPic.Draw_YDA_111(arrHigh(), arrVAL(), arrLow(), strPicPath, str标本号)
                    If strPic <> "" Then
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & strPic
                    End If
                End If
            End If
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function

Private Sub ReadConf(strIniFile As String, ByVal strItems As String, ByRef strReturn As String)
    '读配置文件
    'strIniFile :ini配置文件名
    'strItems   :要读取的字段，如：DATABASE NAME=th_xlb中的 DATABASE NAME,多个字段以,分隔
    'strReturn  :读取的返回值，如:DATABASE NAME=th_xlb中的 th_xlb ，多个反回值以,分隔
    
    Dim arrLine
    Dim strLine As String

    Dim blnOpen As Boolean
    Dim objFilesys As New FileSystemObject
    Dim objStream As TextStream

    Dim intLoop As Long, intCount As Integer
    Dim arrItems As Variant
    
    On Error GoTo ErrHand
    
    '从配置文件中读取保存在文件中的配置cotf
    intCount = UBound(Split(strItems, ","))
    ReDim arrItems(intCount, 2)
    For intLoop = 0 To intCount
        arrItems(intLoop, 0) = Split(strItems, ",")(intLoop)
    Next
    
    If objFilesys.FileExists(strIniFile) Then
        Set objStream = objFilesys.OpenTextFile(strIniFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            
            arrLine = Split(strLine, "=")
            If UBound(arrLine) = 1 Then
            
                For intLoop = 0 To intCount
                    If Trim(arrLine(0)) = Trim(arrItems(intLoop, 0)) Then
                        arrItems(intLoop, 1) = Trim(arrLine(1))
                        Exit For
                    End If
                Next
           End If
        Loop
        objStream.Close
        strReturn = ""
        For intLoop = 0 To intCount
            strReturn = strReturn & "," & arrItems(intLoop, 1)
        Next
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        blnOpen = False
    End If

    Exit Sub
ErrHand:
    If Error = 1 Then
        Resume
    End If
    If blnOpen Then objStream.Close
End Sub

'高斯消元法解三元一次方程
Public Sub GetResult(arrPara() As Double)
    Dim arrA(2, 3) As Double
    Dim intRows As Integer: intRows = 3
    Dim intL As Integer: intL = intRows - 1
    Dim i As Integer, j As Integer, l As Integer, n As Integer, m As Integer, k As Integer
    Dim arrTemp(2) As Double
    
    arrA(0, 0) = 1 / 120 ^ 2: arrA(0, 1) = 1 / 120: arrA(0, 2) = 1: arrA(0, 3) = arrPara(0)
    arrA(1, 0) = 1 / 70 ^ 2: arrA(1, 1) = 1 / 70: arrA(1, 2) = 1: arrA(1, 3) = arrPara(1)
    arrA(2, 0) = 1 / 30 ^ 2: arrA(2, 1) = 1 / 30: arrA(2, 2) = 1: arrA(2, 3) = arrPara(2)

    Do
        n = 0
        For l = k To intL - 1
            arrTemp(n) = arrA(l + 1, k) / arrA(k, k)
            n = n + 1
        Next
        
        For i = k + 1 To intRows - 1
            For j = k To intRows
                arrA(i, j) = arrA(i, j) - arrTemp(m) * arrA(k, j)
            Next
            m = m + 1
        Next
        m = 0
        
        k = k + 1
    Loop While (k < intRows)
    k = l - 1

    Do
        n = 0
        For l = k To 0 Step -1
            arrTemp(n) = arrA(k - l, k + 1) / arrA(k + 1, k + 1)
            n = n + 1
        Next
        
        m = 0
        For i = k To 0 Step -1
            For j = k To intRows
                arrA(k - i, j) = arrA(k - i, j) - arrTemp(m) * arrA(k + 1, j)
            Next
            m = m + 1
        Next

        k = k - 1
    Loop While (k >= 0)
  
    For i = 0 To intRows - 1
        arrPara(i) = arrA(i, intRows) / arrA(i, i)
    Next
End Sub


