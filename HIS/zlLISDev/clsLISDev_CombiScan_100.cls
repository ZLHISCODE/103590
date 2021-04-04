VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CombiScan_100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:COMBI SCAN 100型尿检仪
'厂商:德国科宝
'作者:曾超
'用户:望江
'修改:2009-07-09 兼容500
Private mblnStart As Boolean '开始
Private mblnEnd As Boolean   '结束
Private mstrData As String '他待解析串

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
    '设备标准接口
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    
    Dim lng_Stx As String, lng_Etx As Long, str_Code As String
    Dim varIn As Variant '存放输入数据
    Dim strLine As String
    Dim str标本号 As String, str日期 As String, str检验人标本质控 As String, str指标 As String, str值 As String
    Dim str要提取指标 As String, var要提取指标 As Variant
    Dim iStrPoint As Integer
    Dim strNow As String
    Dim aItem() As String
    Dim strData As String
    Dim intLoop As Integer
    Dim strTmp As String
    Dim aData() As String
    Dim intCount As Integer
    
    If InStr(strSource, "URINALYSIS") = 0 Or InStr(strSource, "----------------------") = 0 Then
        strReserved = strSource
        Exit Sub
    End If
    
    strNow = Mid(strSource, InStr(strSource, "URINALYSIS") + Len("URINALYSIS"), InStr(strSource, "--------------------") - Len("------------------"))
    strNow = Replace(strNow, "-", "")
    strReserved = Mid(strSource, InStr(strSource, "----------------------") + Len("----------------------"))
    aItem = Split(strNow, vbCr)
    
    
    For intLoop = 0 To UBound(aItem)
        If InStr(aItem(intLoop), "CombiScan") > 0 And strData = "" Then
            '时间
            aData = Split(DelSapce(aItem(intLoop + 1)), " ")
            strData = strData & "||" & aData(0) & " " & aData(1)
            '标本号
            aData = Split(DelSapce(aItem(intLoop + 2)), " ")
            strData = strData & "|" & Val(aData(1)) & "| |尿|0"
            intCount = intLoop + 4
        End If
        
        If intLoop >= intCount Then
            '填写结果
            aData = Split(DelSapce(Trim(aItem(intLoop))), " ")
            If UBound(aData) > 0 Then
                strData = strData & "|" & Replace(aData(0), "*", "") & "|" & Replace(aData(1), "neg", "阴性")
            End If
        End If
    Next
    
    strResult = Mid(strData, 3)
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    SendSample = ""
End Function

Private Function DelSapce(strLine As String) As String
    '功能       删除多余的空格
    Dim intLoop  As Integer
    Dim strNow As String
    strNow = strLine
    For intLoop = 20 To 0 Step -1
        strNow = Replace(strNow, Space(intLoop), Space(1))
    Next
    DelSapce = strNow
End Function
'------------------------------------------------------------------------------------------------------------
'    Dim iStrPoint As Integer '字符指针
'    Dim strBuffer As String, strData As String
'    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
'
'    Dim varIn As Variant '存放输入数据
'    Dim strTime As String
'
'    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
'    Dim strLine As String
'    Dim str标本号 As String, str日期 As String, str检验人标本质控 As String, str指标 As String, str值 As String
'    Dim str要提取指标 As String, var要提取指标 As Variant
'
'    On Error GoTo errHandle
'    str检验人标本质控 = " |血液|0"
'    '初始化返回值
'    strReserved = "": strResult = "": strCmd = ""
'    '-- 定义要提取的指标
'    str要提取指标 = "Temp;FIO2;pH;pCO2;pO2;HCO3-;sO2;tCO2;ABE;SBC;SBE"
'    var要提取指标 = Split(str要提取指标, ";")
'
'    If strSource = "" Then Exit Sub
'    Call WriteLog("Abl5.Analyse 进入 ", strSource, mstrData)
'    varIn = Split(strSource, vbNewLine)
'    If UBound(varIn) <= 0 Then
'        varIn = Split(strSource, Chr(10))
'    End If
'
'    If mblnStart = True And mblnEnd = False Then
'            'strReserved = strSource
'        mstrData = mstrData & strSource
'        Call WriteLog("Abl5.Analyse 收取数据 ", strSource, mstrData)
'        If mstrData Like "*Analyzer *" Then
'            mblnEnd = True
'            Call WriteLog("Abl5.Analyse 找到结束符 ", strSource, mstrData)
'        End If
'    End If
'
'
'    If strSource Like "*ABL5*" Then
'        mblnStart = True
'        'strReserved = strSource
'        Call WriteLog("Abl5.Analyse 找到开始符 ", strSource, mstrData)
'        If strSource Like "*Analyzer *" Then
'            mblnEnd = True
'            Call WriteLog("Abl5.Analyse 全部传入,找到结束符 ", strSource, mstrData)
'        End If
'        mstrData = strSource
'    End If
'
'    If mblnStart And mblnEnd And mstrData <> "" Then
'        Call WriteLog("Abl5.Analyse 开始解码 ", strSource, mstrData)
'    End If
'    Exit Sub
'errHandle:
'    Call WriteLog("Abl5.Analyse 解码错误", strSource, strResult & vbNewLine & Err.Description)




