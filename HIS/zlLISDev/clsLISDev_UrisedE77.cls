VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_UrisedE77"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Urised E77 (尿沉渣分析仪)
'厂商:未知
'作者:陈东
'用户:

Private mstrCode As String

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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
        
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号^是否急诊^条码号  可以只有样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
   
    Dim lng_S As Long, lng_E As Long
    Dim str_Code As String '一个完整记录
    Dim str_Line As String '一行数据
    Dim str_标本号 As String, str_项目 As String, str_结果 As String, str_日期 As String, str_Items As String
    Dim var_Code As Variant, i As Integer
    
    Dim objFile As New FileSystemObject
    
    Dim strImgPath As String
    Dim strFilename As String, strImg As String
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Else
        lng_E = InStr(mstrCode, Chr(&H3))
    End If
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '-------------------------------------------------------------
        var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
        str_标本号 = "": str_日期 = "": str_项目 = "": str_结果 = "": str_Items = ""
        For i = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(i)
            If str_Line Like "Sample ID :*" Then
                str_标本号 = Mid(str_Line, 41)
            ElseIf str_Line Like "Date&Time :*" Then
                str_日期 = Format(Now, "yyyy-MM-dd HH:mm")
            ElseIf InStr(",BIL,UBG,KET,ASC,GLU,PRO,BLD,PH ,NIT,LEU,SG.,", "," & Left(str_Line, 3) & ",") > 0 Then
                str_项目 = Trim(Left(str_Line, 3))
                str_结果 = Trim(Mid(str_Line, 61, 20))
                If str_结果 <> "" And str_项目 <> "" Then
                    str_Items = str_Items & "|" & str_项目 & "|" & str_结果
                End If
            ElseIf InStr(",RBC,.RBC,.RBCi,.RBCd,.RBC-G1,.RBC-oth,WBC,.WBC,.WBCc,CRY,.CRY,.CaOx,.CaOxm,.CaOxd,.TRI,.URI,.CaPh,.AMO,.CYS,.LEU,.TYR,.ATY,HYA,PAT,.PAT,.C-HGR,.C-GRA,.C-NEC,.C-RBC,.C-WBC,.C-CRY,.C-MIC,.C-FAT,.C-WAX,.C-MIX,NEC,.NEC,.s-TRA,.d-TRA,.REN,EPI,YEA,BAC,.BAC,.BACr,.BACc,MUC,SPRM,LIP,.LIP,.REN-L,.LDR,.CHOL,TRV,SCH,ART,", "," & Trim(Left(str_Line, 18)) & ",") > 0 Then
                str_项目 = Trim(Left(str_Line, 18))
                str_结果 = Replace(Trim(Mid(str_Line, 21, 20)), ",", ".")
                
                If str_结果 <> "" And str_项目 <> "" Then
                    str_Items = str_Items & "|" & str_项目 & "|" & str_结果
                End If
            End If
        Next
        If Val(str_标本号) > 0 And str_日期 <> "" And str_Items <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |尿液|0" & str_Items
            Call ReadConf(App.Path & "\UrisedE77.ini", "IMG_PATH", strImgPath)
            strImgPath = Trim(strImgPath)
            If strImgPath <> "" Then
                
                If objFile.FolderExists(Trim(strImgPath)) = True Then
                    strFilename = Dir(strImgPath & "\" & str_标本号 & "_*_*_*_*_*_*.JPG")      ' 找寻第一项。
                     
                    strImg = ""
                    
                    Do While strFilename <> ""   ' 开始循环。
                        '9074_090421_02_59_11_-_01.JPG
                        strImg = strImg & "^" & Replace(UCase(Split(strFilename, "_")(6)), ".JPG", "") & ";110;1;" & strImgPath & "\" & strFilename
                        strFilename = Dir   ' 查找下一个文件。
                    Loop
                    If strImg <> "" Then
                        strImg = Mid(strImg, 2)
                        If strResult <> "" Then strResult = strResult & vbCrLf & strImg
                    End If
                End If
            End If
        End If
        '-------------------------------------------------------------
        lng_S = InStr(mstrCode, Chr(&H2))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        Else
            lng_E = InStr(mstrCode, Chr(&H3))
        End If
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)

End Sub

Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5个元素：条码
    '   第6～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
        SendSample = ""
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
                    If Trim(arrLine(0)) = arrItems(intLoop, 0) Then
                        arrItems(intLoop, 1) = arrLine(1)
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
