VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_KY90"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:尿沉渣
'医院:淮安三院尿沉渣
'开发:陈玉强
'时间:2015/1/17
'厂家:


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

    Dim aResults() As String
    Dim cnsql As New ADODB.Connection
    Dim cnstring As String
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsItem As ADODB.Recordset
    Dim lngResultCount As Long
    Dim blnUnion As Boolean
    Dim lngSampleID As Long
    Dim blnNew As Boolean           '是否新标本号
    Dim lngSampleNumber As Long     '标本ID
    Dim lngloop As Long
    Dim intLoop As Integer
    Dim lng_count As Long
    Dim st日期 As String, en日期 As String, str日期 As String
    Dim strSer As String, strDB As String, strPWD As String
    Dim strValues As String, varValues As Variant
    Dim strItemValue As String, i As Integer
    Dim str标本号 As String
    Dim str结果 As String


    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    lng_count = -1
    ResultFromFile = aResults

    With cnsql
        Set cnsql = New ADODB.Connection
        If cnsql.State = adStateOpen Then cnsql.Close
        On Error Resume Next
        cnstring = "DSN=zllis;UID=root;PSW=1234"    'dsn系统odbc,uid 用户名,psw密码
        cnsql.Open cnstring
        st日期 = Format(dtStart, "yyyy-mm-dd")
        en日期 = Format(dtEnd, "yyyy-mm-dd")
        If lngSampleNO > 0 Then
            strSQL = "SELECT a.date,a.SampleCode,a.RedCell ,a.WhiteCell ,a.EpithelialCell,a.MucousSilk,a.Crystallize,a.UricAcidCrystal,a.Trichomonad,a.Fungus,a.Sperm,a.HyalineCast," & _
                   " a.RedCellCast,a.WhiteCellCast,a.WaxyCast,a.ThinGrainCast,a.ThickGrainCast,a.CircleEpi,a.CellBlock,a.unnormal_red_cell,a.unknow_cell," & _
                   " b.colour,b.sharpness,b.glucose,b.bilirubin,b.ketone,b.propotion,b.PHvalue,b.protein,b.URO,b.nitrite,b.occultblood,b.Whitecell as bWhitecell," & _
                   " b.vitaminC , b.calcium, b.creatinine, b.albumin " & _
                   " from nzjyjg a,drychemical b " & _
                   " where a.Date=b.Date and a.SampleCode=b.SampleCode and a.Date between '" & st日期 & "' and '" & en日期 & "' and a.SampleCode=" & lngSampleNO
        Else
            strSQL = "SELECT a.date,a.SampleCode,a.RedCell ,a.WhiteCell ,a.EpithelialCell,a.MucousSilk,a.Crystallize,a.UricAcidCrystal,a.Trichomonad,a.Fungus,a.Sperm,a.HyalineCast," & _
                   " a.RedCellCast,a.WhiteCellCast,a.WaxyCast,a.ThinGrainCast,a.ThickGrainCast,a.CircleEpi,a.CellBlock,a.unnormal_red_cell,a.unknow_cell," & _
                   " b.colour,b.sharpness,b.glucose,b.bilirubin,b.ketone,b.propotion,b.PHvalue,b.protein,b.URO,b.nitrite,b.occultblood,b.Whitecell as bWhitecell," & _
                   " b.vitaminC , b.calcium, b.creatinine, b.albumin " & _
                   " from nzjyjg a,drychemical b " & _
                   " where a.Date=b.Date and a.SampleCode=b.SampleCode and a.Date between '" & st日期 & "' and '" & en日期 & "'"
        End If

        Set rsTmp = cnsql.Execute(strSQL)
        Do Until rsTmp.EOF
            strItemValue = ""
            str标本号 = Val(rsTmp!SampleCode)
            str日期 = Format(rsTmp!date, "YYYY-MM-DD")
            intLoop = rsTmp.Fields.count - 1
            For i = 2 To intLoop
                strItemValue = strItemValue & "|" & i & "|" & Replace(Replace(Replace(Nvl(rsTmp.Fields(i), " "), vbCrLf, ""), vbCr, ""), "|", "")
            Next
            If strItemValue <> "" Then
                If aResults(UBound(aResults)) <> "" Then
                    ReDim Preserve aResults(UBound(aResults) + 1)
                End If
                str结果 = str日期 & "|" & str标本号 & "| |尿液|0" & strItemValue
                aResults(UBound(aResults)) = str结果
            End If
            strItemValue = ""
            rsTmp.MoveNext
        Loop
        ResultFromFile = aResults
    End With
    '    ResultFromFile = aResults
End Function

Public Sub AnalyConf(strItem As String, strCO As String)
    Dim arrLine
    Dim strLine As String
    Dim strFile As String
    Dim blnOpen As Boolean
    Dim objFilesys As New FileSystemObject
    Dim objStream As textStream
    On Error GoTo ErrHand
    Dim intLoop As Long


    '从配置文件中读取保存在文件中的配置cotf
    strFile = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "U2_Config.ini"
    If objFilesys.FileExists(strFile) Then
        Set objStream = objFilesys.OpenTextFile(strFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            If strLine = "" Then Exit Do
            arrLine = Split(strLine, "=")
            If arrLine(0) = strItem Then
                strCO = arrLine(1)
                Exit Sub
            End If
        Loop
        objStream.Close
        blnOpen = False
    End If

    Exit Sub
ErrHand:
    If Error = 1 Then
        Resume
    End If
    If blnOpen Then objStream.Close
End Sub

Public Function Fun_结果(ByVal StrInput As String, ByVal strKey As String) As String
    Dim str结果 As String
    Dim strVar As Variant
    Dim i As Long
    Fun_结果 = ""
    If StrInput = "" Then Exit Function
    StrInput = Replace(Replace(Replace(StrInput, Chr(10), ""), Chr(13), ""), "<p", "p")
    strVar = Split(StrInput, ";")
    For i = LBound(strVar) To UBound(strVar)
        If Split(Replace(strVar(i), vbCrLf, ""), "=")(0) = strKey Then
            Fun_结果 = Trim(Split(strVar(i), "=")(1))
            Exit Function
        End If
    Next i
End Function










