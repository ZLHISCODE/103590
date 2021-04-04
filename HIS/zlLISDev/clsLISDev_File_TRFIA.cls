VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_TRFIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:广州丰华
'厂商:
'作者:白晓艳
'用户:
'时间:20151223
Public objFso As New FileSystemObject

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

   Dim aResults() As String, objFile As New Scripting.FileSystemObject
   Dim strChenel As String, str操作员 As String, str日期 As String
   Dim strFilePath As String, strFilename As String, k As Integer, int质控数 As Integer, strRQ As String
   Dim strNeedFile As String  '确定现在需要提取的标本的日期
   Dim FILE开始日期 As String, FILE结束日期 As String, strTemp As String
   
   Dim objStream As textStream, lngResultCount As Long
   Dim strLine As String, varItem As Variant, varTmp As Variant
   Dim str标本号 As Long, blnReturn As Boolean
   Dim lngRow As Long, varLine As Variant
   Dim tempstr As String, NowFile As String, NewFile As String, str检验结果 As String

  
   '避免出现下标错误
    ResultFromFile = aResults
    ReDim aResults(0)
    
    NewFile = strFile
    Set objStream = objFile.OpenTextFile(NewFile, ForReading, False, -2)
    lngRow = 0
    str日期 = ""
    str检验结果 = ""
    str标本号 = 0


    FILE开始日期 = Year(dtStart) & IIf(Len(Month(dtStart)) = 1, 0 & Month(dtStart), _
                   Month(dtStart)) & IIf(Len(Day(dtStart)) = 1, 0 & Day(dtStart), Day(dtStart))
    FILE结束日期 = Year(dtEnd) & IIf(Len(Month(dtEnd)) = 1, 0 & Month(dtEnd), _
                   Month(dtEnd)) & IIf(Len(Day(dtEnd)) = 1, 0 & Day(dtEnd), Day(dtEnd))

              
  ''数据库格式解析
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim str As String
    
    On Error GoTo errHand
    
    cn.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & strFile '连接数据库    D:\Uranus\ Modules\丰华时间分辨\Data.dat
    If lngSampleNO = -1 Then
         rs.Open "select a.A17 as A17 ,a.A2 as A2 ,b.C4 as  C4,b.C6 as C6 from h_Patient a inner join h_Result b on a.A1=b.C2 where len(b.C4)<>0   and len(b.C6)<>0   and  a.A17  between '" & FILE开始日期 & "' and '" & FILE结束日期 & "'", cn, adOpenKeyset, adLockOptimistic '连接数据表"
    Else
         rs.Open "select a.A17 as A17 ,a.A2 as A2 ,b.C4 as  C4,b.C6 as C6 from h_Patient a inner join h_Result b on a.A1=b.C2 where len(b.C4)<>0   and  len(b.C6)<>0   and b.C0= " & lngSampleNO & " and a.A17 = '" & FILE开始日期 & "' ", cn, adOpenKeyset, adLockOptimistic '连接数据表"
    End If
    If rs.RecordCount > 0 Then
         rs.MoveFirst
    End If
    
            Do While rs.EOF = False
             
               str日期 = rs.Fields("A17")
               str日期 = Mid(str日期, 1, 4) + "-" + Mid(str日期, 5, 2) + "-" + Mid(str日期, 7, 2)
'               str日期 = Format(str日期, "YYYY-MM-DD")
               str标本号 = Val(rs.Fields("A2"))
               '传染8项  Trim(rs.Fields("C6") & " ")  IIf(Len(rs.Fields("C6")) = 0, " ", rs.Fields("C6")) Trim(rs.Fields("C4") & " ")
               str检验结果 = "|" + rs.Fields("C4") + "|" + rs.Fields("C6")

             If str检验结果 <> "" And str标本号 > 0 And str日期 <> "" Then
                   If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                   aResults(UBound(aResults)) = str日期 & "|" & Val(str标本号) & "| |" & "血清|0" & str检验结果
                   Debug.Print aResults(UBound(aResults))
                   str检验结果 = ""
                   str日期 = ""
                   str标本号 = 0
                End If

                rs.MoveNext
             Loop
        cn.Close
 
    ResultFromFile = aResults
    
    Exit Function
errHand:
    WriteLog "ResultFromFile", "", Err.Number & " " & Err.Description
End Function



