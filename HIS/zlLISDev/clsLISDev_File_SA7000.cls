VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_SA7000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:SA_7000
'厂商:
'作者:白晓艳
'用户:旬阳县医院
'时间:20151217
 
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
   Dim FILE开始日期 As Date, FILE结束日期 As Date, strTemp As String
   
   Dim objStream As textStream, lngResultCount As Long
   Dim strLine As String, varItem As Variant, varTmp As Variant
   Dim str标本号 As Long, blnReturn As Boolean
   Dim lngRow As Long, varLine As Variant
   Dim tempstr As String, NowFile As String, NewFile As String, str项目 As String, str检验结果 As String

  


   '避免出现下标错误
    ResultFromFile = aResults
    ReDim aResults(0)
    
    NewFile = strFile
    Set objStream = objFile.OpenTextFile(NewFile, ForReading, False, -2)
    lngRow = 0
    str日期 = ""
    str项目 = ""
    str检验结果 = ""
    str标本号 = 0
            
  ''数据库格式解析
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim str As String
    
    On Error GoTo errHand
    
    cn.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & strFile '连接数据库
    If lngSampleNO = -1 Then
          rs.Open "select * from TestData where 日期 between Cdate('" & Format(dtStart, "yyyy-MM-dd") & "') and  Cdate('" & Format(dtEnd, "yyyy-MM-dd") & "')", cn, adOpenKeyset, adLockOptimistic  '连接数据表
    Else
          rs.Open "select * from TestData where 日期 = Cdate('" & Format(dtStart, "yyyy-MM-dd") & "') and TestDataID=" & lngSampleNO, cn, adOpenKeyset, adLockOptimistic  '连接数据表"
    End If
    If rs.RecordCount > 0 Then
         rs.MoveFirst
    End If

           Do While rs.EOF = False
           
               str日期 = Format(rs.Fields("日期") & " " & rs.Fields("时间"), "yyyy-mm-dd  hh:mm:ss")
               str标本号 = Val(rs.Fields("TestDataID"))

'                当前结果
              str检验结果 = "|血浆粘度|" + Trim(rs.Fields("血浆粘度") & "") + "|血沉|" + Trim(rs.Fields("血沉") & "") + "|压积|" + Trim(Format(rs.Fields("压积"), "0.00") & "") + "|全血低切相对指数|" + _
                             Trim(rs.Fields("全血低切相对指数") & "") + "|全血高切相对指数|" + Trim(rs.Fields("全血高切相对指数") & "") + "|血沉方程K值|" + Trim(rs.Fields("血沉方程K值") & "") + _
                             "|红细胞聚集指数|" + Trim(rs.Fields("红细胞聚集指数") & "") + "|全血低切还原粘度|" + Trim(rs.Fields("全血低切还原粘度") & "") + "|全血高切还原粘度|" + _
                             Trim(rs.Fields("全血高切还原粘度") & "") + "|红细胞刚性指数|" + Trim(rs.Fields("红细胞刚性指数") & "") + "|红细胞变形指数TK|" + Trim(Format(rs.Fields("红细胞变形指数TK"), "0.00") & "")

                 If str检验结果 <> "" And str标本号 <> 0 And str日期 <> "" Then
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




