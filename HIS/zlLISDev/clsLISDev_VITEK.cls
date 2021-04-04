VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_VITEK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:  微生物 VITEK
'医院:  晋煤,梅河口医院
'作者:  曾超

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
    '补充图像的方式：
    '                   1.图像数据跟随指标数据后，使用回车换行符来分隔。
    '                   2.有多个图像数据时使用"^"来分隔
    '                   3.单个图像数据格式: 图像画法 0=直方图  1=散点图
    '                     a) 直方图: 图像名称;图像画法(0=直方图  1=散点图);X1;X2;X3;X4;X5...
    '                     b) 散点图: 图像名称;图像画法(0=直方图  1=散点图):
    '                        例:00000100001000010000100010;00000100001000010000100010;
    '                        说明:1.散点图以点阵方式保存每一行使用分号来分隔.
    '                             2.有多少个分号就有多少行
    '                             3.每一行有多少个点由每一行的长度来确定
    '                             3.画图的方向是从最上边向下画，如有65*65的图就是从65行开始画(最上边开始画)
    

    Dim strData As String
    Dim aData() As String
    Dim lngloop As Long
    Dim lng_S As Long, lng_E As Long
    Dim strNow As String
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
'
'    If InStr(Replace(strSource, mstrLastInput, ""), Chr(&H5)) > 0 Then
'        strCmd = Chr(&H6)
'    End If
'
'    If InStr(Replace(strSource, mstrLastInput, ""), Chr(&H4)) > 0 Then
'        strCmd = Chr(&H6)
'    End If
'
'    If InStr(Replace(strSource, mstrLastInput, ""), Chr(&H3)) > 0 Then
'        strCmd = Chr(&H6)
'    End If
'    If InStr(Replace(strSource, mstrLastInput, ""), Chr(&H1E)) > 0 Then
'        strCmd = Chr(&H6)
'    End If
'    If InStr(Replace(strSource, mstrLastInput, ""), Chr(&H1D)) > 0 Then
'        strCmd = Chr(&H6)
'    End If
'    '没有找到开始位置时退出
'    If Not strSource Like "**mtrsl|iiV2|is*|itSU*" Then
'        strReserved = strSource
'        mstrLastInput = strReserved
'        Exit Sub
'    End If
'
'    '没有找到位置位置时退出
'    lngStart = InStr(strSource, "mtrsl|iiV2|is")
'    If InStr(lngStart, strSource, "") <= 0 Then
'        strReserved = strSource
'        mstrLastInput = strReserved
'        Exit Sub
'    End If
'    lngEnd = InStr(lngStart, strSource, "") + 2
    '---------------------------------------------------------
    If InStr(strSource, Chr(&H5)) > 0 Then strCmd = Chr(&H6)
    If InStr(strSource, Chr(&H4)) > 0 Then strCmd = Chr(&H6)
    If InStr(strSource, Chr(&H3)) > 0 Then strCmd = Chr(&H6)
    If InStr(strSource, Chr(&H1E)) > 0 Then strCmd = Chr(&H6)
    If InStr(strSource, Chr(&H1D)) > 0 Then strCmd = Chr(&H6)
    
    mstrCode = mstrCode & Replace(strSource, Chr(&H1E), "")
    
    lng_S = InStr(mstrCode, Chr(&H5))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H4))
    Else
        lng_E = InStr(mstrCode, Chr(&H4))
    End If
    
    Do While lng_S > 0 And lng_E > 0
        strNow = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        If strNow Like "**mtrsl|iiV2|is*|itSU*" Or strNow Like "**mtrsl|iiV2|is*|itID*" Then
            aData = Split(strNow, "|")
            strData = Format(Mid(aData(7), 3), "yyyy-mm-dd") & "|" & Val(Mid(aData(9), 3)) & "| |ABCDEFGHI|0"
            
            '解析结果
            For lngloop = 7 To UBound(aData)
                
                If Mid(aData(lngloop), 1, 2) = "o1" Then
                    '得到细菌
                    strData = Replace(strData, "ABCDEFGHI", Mid(aData(lngloop), 3))
                ElseIf Mid(aData(lngloop), 1, 2) = "a1" Then
                    '得到结果
                    If InStr(",R,S,I,", "," & Trim(UCase(Mid(aData(lngloop + 3), 3))) & ",") > 0 Then
                        strData = strData & "|" & Trim(Mid(aData(lngloop), 3)) & "|" & Trim(UCase(Mid(aData(lngloop + 3), 3))) & "^1^" & Trim(Mid(aData(lngloop + 2), 3))
                    ElseIf InStr(",+,-,", "," & Trim(UCase(Mid(aData(lngloop + 3), 3))) & ",") > 0 Then
                        strData = strData & "|" & Trim(Mid(aData(lngloop), 3)) & "|" & "^^" & Trim(Mid(aData(lngloop + 3), 3))
                    End If
                End If
            Next
            If Trim(strData) <> "" Then
                strResult = strResult & "||" & strData
            End If
            
        End If
        lng_S = InStr(mstrCode, Chr(&H5))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H4))
        Else
            lng_E = InStr(mstrCode, Chr(&H4))
        End If
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
End Sub



Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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
    Dim varItem As Variant, strItem As String, i As Integer, strSampleNO As String
    blnSuccess = True
   
End Function










