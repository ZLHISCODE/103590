VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_FB01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:粪便分析仪器
'作者:李骏驰
'用户:宾川县妇幼保健院

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
    Dim str标本号 As String, str项目 As String
    
    Dim rsErs As ADODB.Recordset
    Dim str日期 As String

    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
        
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "SELECT a.Bianhao,a.riqi,a.Fenbianyanse,a.Xingzhuang,a.Hongxibao,a.Baixibao,a.Shiwuxibao,a.Buxiaohuashiwu,a.Gouchongluan,a.Yuanchong,a.Baonang" & _
                ",a.Huichongluan ,a.Qita,a.Bianchongluan,a.Jushixibao,a.Yinxiehuaxue,a.Qitachongluan,a.Jiadishi,a.Yinxiemianyi,a.lujing  " & _
                ",a.Changdichong ,a.Xialeidengshi,a.Xianbingdu,a.Lunzhuanbingdu,a.Kesaqi,a.Youmenluoxuan" & _
                " FROM report as A Where A.riqi >= Cdate('" & Format(dtStart, "yyyy-MM-dd") & _
                "') And A.riqi< Cdate('" & Format(DateAdd("d", 1, dtEnd), "yyyy-MM-dd") & "') "
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        ReDim aResults(0)
        Do While Not rsTmp.EOF
                str日期 = Format("" & rsTmp("riqi"), "yyyy-MM-dd")
                str标本号 = Val("" & rsTmp("Bianhao"))

                str项目 = ""
                str项目 = "|Fenbianyanse|" & Nvl(Trim(rsTmp("Fenbianyanse")), " ") & "|Xingzhuang|" & Nvl(Trim(rsTmp("Xingzhuang")), " ") & "|Hongxibao|" & Nvl(Trim(rsTmp("Hongxibao")), " ") & "|Baixibao|" & Nvl(Trim(rsTmp("Baixibao")), " ") & _
                          "|Shiwuxibao|" & Nvl(Trim(rsTmp("Shiwuxibao")), " ") & "|Buxiaohuashiwu|" & Nvl(Trim(rsTmp("Buxiaohuashiwu")), " ") & "|Gouchongluan|" & Nvl(Trim(rsTmp("Gouchongluan")), " ") & "|Yuanchong|" & Nvl(Trim(rsTmp("Yuanchong")), " ") & _
                          "|Baonang|" & Nvl(Trim(rsTmp("Baonang")), " ") & "|Huichongluan|" & Nvl(Trim(rsTmp("Huichongluan")), " ") & "|Qita|" & Nvl(Trim(rsTmp("Qita")), " ") & "|Bianchongluan|" & Nvl(Trim(rsTmp("Bianchongluan")), " ") & "|Jushixibao|" & Nvl(Trim(rsTmp("Jushixibao")), " ") & _
                          "|Yinxiehuaxue|" & Nvl(Trim(rsTmp("Yinxiehuaxue")), " ") & "|Qitachongluan|" & Nvl(Trim(rsTmp("Qitachongluan")), " ") & "|Jiadishi|" & Nvl(Trim(rsTmp("Jiadishi")), " ") & "|Yinxiemianyi|" & Nvl(Trim(rsTmp("Yinxiemianyi")), " ") & "|Changdichong|" & Nvl(Trim(rsTmp("Changdichong")), " ") & _
                          "|Xialeidengshi|" & Nvl(Trim(rsTmp("Xialeidengshi")), " ") & "|Xianbingdu|" & Nvl(Trim(rsTmp("Xianbingdu")), " ") & "|Lunzhuanbingdu|" & Nvl(Trim(rsTmp("Lunzhuanbingdu")), " ") & "|Kesaqi|" & Nvl(Trim(rsTmp("Kesaqi")), " ") & "|Youmenluoxuan|" & Nvl(Trim(rsTmp("Youmenluoxuan")), " ")
                If str项目 <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                        aResults(UBound(aResults)) = str日期
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val("" & str标本号)             '样本序号
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "                             '检验人
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "血液"                          '标本
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"                                  '质控
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & str项目                               '项目和结果
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & _
                            "JB;103;0;" & Nvl(Trim(rsTmp("lujing")), " ") + "\jinbiao_01.jpg" & "^XWJ;103;0;" & Nvl(Trim(rsTmp("lujing")), " ") + "\xianweijing_01_10.jpg" & "^XZ;103;0;" & Nvl(Trim(rsTmp("lujing")), " ") + "\xingzhuang_01.jpg"
                End If

            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function

