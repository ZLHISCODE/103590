VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_CAPILLARYS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:CAPILLARYS 电永仪
'厂商:Sebia
'作者:陈东
'用户:烟台山医院
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
    
        Dim aResults() As String
        Dim cnAccess As New ADODB.Connection
        Dim strFilePath As String, objFile As New Scripting.FileSystemObject
        Dim strSQL As String, rsTmp As ADODB.Recordset
        Dim lngResultCount As Long, strBh As String
        Dim blnUnion As Boolean, strimg As String, iField As Integer, i As Integer
    
        Dim cnErs As New ADODB.Connection
        Dim strFileErs As String, rsErs As ADODB.Recordset
        Dim strImgErs As String, str日期 As String
        Dim strChar As String
        Dim intMaxValue As Integer
        On Error GoTo hErr
100     strFilePath = objFile.GetParentFolderName(strFile)
102     If Len(strFilePath) = 0 Then strFilePath = App.Path
    
104     strChar = Replace(strFile, strFilePath & "\", "")
106     strChar = Mid(strChar, 4, 1)
    
108     strFileErs = strFile
    
110     If Trim(Dir(strFileErs)) <> "" Then
112         With cnErs
114         .ConnectionString = "DBQ=" & strFileErs & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
116         .Open
118         If .State = adStateClosed Then strFileErs = ""
            End With
        Else
120         strFileErs = ""
        End If
        
122     lngResultCount = -1
   
124     strImgErs = ""
126     WriteLog "开始!", "", strFileErs
128     If strFileErs <> "" Then
130         strSQL = "SELECT val(seq) As 顺序号,programma as 类型, data_analisi as 日期, " & _
                     "A.fraz_1,fraz_2,fraz_3,fraz_4,fraz_5,fraz_6,fraz_7,fraz_8,fraz_9,fraz_10,rapp1,Curva " & _
                     " FROM Anagrafica AS A Where Format(A.data_analisi,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "' And Format(A.data_analisi,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & _
                     "' " & IIf(lngSampleNO = -1, "", " And val(seq)=" & lngSampleNO)
132         Set rsErs = cnErs.Execute(strSQL)
134         'WriteLog "执行Sql", rsErs.EOF, strSQL
136         Do Until rsErs.EOF

140             If Val("" & rsErs("顺序号")) <> 0 Then
142                 lngResultCount = lngResultCount + 1
144                 ReDim Preserve aResults(lngResultCount)
                
146                 aResults(lngResultCount) = Format(rsErs("日期"), "yyyy-MM-dd")  '检验时间
                
148                 If "" & rsErs!类型 = "S" Then
150                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                            Val("" & rsErs("顺序号"))   '样本序号 血液标本+100
                    Else
152                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                            Val("" & rsErs("顺序号")) + 200 '样本序号
                    End If
154                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                          " "  '检验人
156                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                          IIf("" & rsErs!类型 = "S", "血液", "尿液")  '标本
                      
158                 aResults(lngResultCount) = aResults(lngResultCount) & "|0"
160                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "01|" & rsErs!fraz_1
162                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "02|" & rsErs!fraz_2
164                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "03|" & rsErs!fraz_3
166                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "04|" & rsErs!fraz_4
168                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "05|" & rsErs!fraz_5
                
170                 If "" & rsErs!类型 = "S" Then
172                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "06|" & rsErs!rapp1
                    Else
174                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "06|" & rsErs!fraz_6
176                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "07|" & rsErs!fraz_7
178                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "08|" & rsErs!fraz_8
180                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "09|" & rsErs!fraz_9
182                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "10|" & rsErs!fraz_10
                    End If
184                 strImgErs = "" & rsErs!Curva
            
186                 If strImgErs <> "" Then
                        Dim maxVal As Long
                    
188                     strImgErs = SebiaImg(Mid(strImgErs, 5, Len(strImgErs) - 8))
190                     maxVal = Val(Split(strImgErs, "|")(0))
192                     If maxVal <= 0 Then maxVal = 3800
194                     strImgErs = "　;0;" & maxVal + 200 & ",300,0-;" & Split(strImgErs, "|")(1)
196                     aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & strImgErs
198                     'WriteLog "电泳", Val("" & rsErs("顺序号")), strImgErs
                    End If
                End If
200             rsErs.MoveNext
            Loop
        End If
202     'WriteLog "结束!", "", strFileErs
204     ResultFromFile = aResults
        Exit Function
hErr:
206     WriteLog "电泳", CStr(Erl()), Err.Description

End Function

Private Function SebiaImg(ByVal strimg As String) As String
'   生成图形数据串
    Dim maxVal As Long
    Do While strimg <> ""
        If maxVal < Val("&H" & Left(strimg, 4)) And Val("&H" & Left(strimg, 4)) < 4000 Then
            maxVal = Val("&H" & Left(strimg, 4))
        End If
        If Val("&H" & Left(strimg, 4)) < 0 Then
            SebiaImg = "1;" & SebiaImg
        ElseIf Val("&H" & Left(strimg, 4)) < 4000 Then
            SebiaImg = CStr(Val("&H" & Left(strimg, 4))) & ";" & SebiaImg
        End If
        strimg = Mid(strimg, 5)
    Loop
    If Right(SebiaImg, 1) = ";" Then SebiaImg = maxVal & "|" & Mid(SebiaImg, 1, Len(SebiaImg) - 1)

End Function


