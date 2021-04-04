VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_MicroScan_AS4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:Microscan AS-4 微生物分析仪
'厂商:美国德灵(DADE BEHRING INC)
'作者:陈东
'用户:南京江北人民医院
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
    
    Dim aResults() As String
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim lngResultCount As Long
    Dim objStream As Scripting.TextStream, strLine As String
    
    Dim str标本号 As String, str药敏项目 As String
    Dim str日期 As String, str细菌 As String, str敏感度 As String
    If Dir(strFile) = "" Then Exit Function
    Set objStream = objFile.OpenTextFile(strFile, ForReading)

    lngResultCount = -1
    ReDim aResults(0)
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        If strLine Like "R,*" Then
            
            ''有两个细菌以上的情况。 先返回上一个细菌结果
            If Val(str标本号) > 0 And str细菌 <> "" And str日期 <> "" And str药敏项目 <> "" Then
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
'                aResults(lngResultCount) = str日期 & "|" & Val(str标本号) & "| |" & str细菌 & "|0" & str药敏项目
                '市三院修改取标本号的后4位
                aResults(lngResultCount) = str日期 & "|" & Val((str标本号)) & "| |" & str细菌 & "|0" & str药敏项目
                str日期 = "": str标本号 = "": str细菌 = "": str药敏项目 = ""
            End If
            
            '一个标本开始,细菌结果
            
            str标本号 = Trim(Replace(Split(strLine, ",")(3), """", ""))
            str标本号 = Mid(str标本号, 1, Len(str标本号) - 1)
            str日期 = Replace(Split(strLine, ",")(6), """", "")
            str日期 = Mid(str日期, 1, 4) & "-" & Mid(str日期, 5, 2) & "-" & Mid(str日期, 7)
            'by陈玉强,有原来的细菌名称修改为细菌编号.原代码取(12)
            str细菌 = Replace(Split(strLine, ",")(11), """", "")
            
            If str日期 >= Format(dtStart, "yyyy-MM-dd") And str日期 <= Format(dtEnd, "yyyy-MM-dd") Then
                
                If lngSampleNO <> -1 Then '要指定标本号
                    If str标本号 <> lngSampleNO Then
                        '标本号不同，不返回
                        str标本号 = ""
                        str日期 = ""
                        str细菌 = ""
                    End If
                End If
            Else
                '日期不符，不返回
                str标本号 = ""
                str日期 = ""
                str细菌 = ""
            End If
            
        ElseIf strLine Like "M,*" Then
            '药敏结果
            If str标本号 <> "" Then
            '修改细菌检验结果,其中r*,blac,esbl 解析成耐药by陈玉强.
                str敏感度 = Replace(Split(strLine, ",")(7), """", "")
                Select Case str敏感度
                       Case "R*": str敏感度 = "R"
                       Case "BLAC": str敏感度 = "R"
                       Case "ESBL": str敏感度 = "R"
                End Select
                str药敏项目 = str药敏项目 & "|" & Replace(Split(strLine, ",")(2), """", "") & "|" & str敏感度 & "^1" & "^" & Replace(Split(strLine, ",")(4), """", "")
'                str药敏项目 = str药敏项目 & "|" & Replace(Split(strLine, ",")(2), """", "") & "|" & Replace(Split(strLine, ",")(7), """", "") & "^1" & "^" & Replace(Split(strLine, ",")(4), """", "")
            
            End If
        ElseIf strLine Like "L,*" Then
            '一个标本结束
            If Val(str标本号) > 0 And str细菌 <> "" And str日期 <> "" And str药敏项目 <> "" Then
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                aResults(lngResultCount) = str日期 & "|" & Val(str标本号) & "| |" & str细菌 & "|0" & str药敏项目
                str日期 = "": str标本号 = "": str细菌 = "": str药敏项目 = ""
            End If
        End If
    Loop
    ResultFromFile = aResults
    
End Function

