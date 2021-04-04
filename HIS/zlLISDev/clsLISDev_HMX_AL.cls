VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HMX_AL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:HMX_AL
'厂商:美国贝克曼库尔特 Beckman coulter
'修改:尹建标
'用户:新疆乌石化医院
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
    '   第1个元素：样本序号 后面加 ^1 表示急诊 后加加 ^条码
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    
    Dim varline As Variant, strItem As String, strItem1 As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String
    Dim blnEnd As Boolean '是否已解析完
    Const strStartChr As String = "01"
    Const strEndChr As String = "" & vbCrLf & vbCrLf & "--------------"
    Dim strNowData As String
    Dim strData As String
    Dim intloop As Integer
    Dim strLineData As String
    Dim aItem As Variant
    Dim strItems As String
    Dim strImage As String
    Dim StrPLt As String
    Dim strDIF1 As String
    Dim frmGraph_Tmp As Form
    Dim lng_E As Long, lng_S As Long
    
    
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    '要做的项目
    strItems = ";WBC;RBC;HGB;HCT;MCV;MCH;MCHC;RDW;PLT;PCT;MPV;PDW;LY#;MO#;NE#;EO#;BA#;LY%;MO%;NE%;EO%;BA%;"
    
    '没有找到开始符和结束符时退出
    mstrCode = mstrCode & strSource
    
    lng_S = InStr(mstrCode, strStartChr)
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, strEndChr)
        If lng_E > 0 Then
            lng_E = InStr(lng_E, mstrCode, "")
        End If
    End If
    
    Do While lng_S > 0 And lng_E > 0
        If lng_S > 0 And lng_E > 0 Then
            '取得当前一段，剩下的返回。下次再解析。
            strNowData = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
            mstrCode = Mid(mstrCode, lng_E + 1)
        End If
    
    
        '开始解析
        varline = Split(strNowData, vbCrLf)
        
        For intloop = 0 To UBound(varline)
            strLine = DelCRC(varline(intloop))
            If InStr(strLine, "DATE") > 0 And UBound(Split(strData, "|")) <= 0 Then
                '日期
                strData = Trim(Mid(strLine, 5))
                strData = "20" & Mid(strData, 7, 2) & "-" & Mid(strData, 1, 2) & "-" & Mid(strData, 4, 2)
            End If
            If InStr(strLine, "TIME") > 0 And UBound(Split(strData, "|")) <= 0 Then
                '时间
                strData = strData & " " & Trim(Mid(strLine, 5))
            End If
            If InStr(strLine, "ID1") > 0 Or InStr(strLine, "ID2") > 0 Then
                '标本号(ID1,ID2)
                If Val(Mid(strLine, 4)) <> 0 And UBound(Split(strData, "|")) <= 0 Then
                    strData = strData & "|" & Val(Mid(strLine, 4)) & "| |血清|0"
                End If
            End If
            If InStr(strLine, "1EC/P") > 0 Then
                '标本号(1EC/P)
                If UBound(Split(strData, "|")) <= 0 Then
                    strData = strData & "|" & Val(Mid(strLine, 6)) & "| |血清|0"
                End If
            End If
            
            If UBound(Split(strLine, " ")) >= 0 Then
                '指标和结果
                If InStr(strItems, ";" & Split(strLine, " ")(0) & ";") > 0 Then
                    strData = strData & "|" & Split(strLine, " ")(0) & "|" & Format(Val(Replace(strLine, Split(strLine, " ")(0), "")), "0.00")
                    
                End If
            End If
            
            '===========================图形=====================================
            If InStr(strLine, "RBCH") > 0 And Len(strLine) >= 100 Then
                strImage = strImage & "^RBC;0;200,262,16-50,66-100,132-200,198-" & Get_ImgData(Trim(Mid(strLine, 6)))
            End If
            
            If Mid(strLine, 1, 1) = "V" And Len(strLine) >= 100 Then
                strImage = strImage & "^V;0;200,262" & Get_ImgData(Mid(strLine, 3))
            End If
            If Mid(strLine, 1, 1) = "C" And Len(strLine) >= 100 Then
                strImage = strImage & "^C;0;200,262" & Get_ImgData(Mid(strLine, 3))
            End If
            If Mid(strLine, 1, 1) = "S" And Len(strLine) >= 100 Then
                strImage = strImage & "^S;0;200,262" & Get_ImgData(Mid(strLine, 3))
            End If
            If InStr(strLine, "PLTF") > 0 And Len(strLine) >= 100 Then
                strImage = strImage & "^PLT;4;200,262,10-2,50-10,100-20,150-30;" & Replace(Mid(Get_ImgData(Mid(strLine, 6)), 2), ";", ",") & "|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0," & StrPLt
            End If
            
            If InStr(strLine, "PLTH") > 0 And Len(strLine) >= 100 Then
                StrPLt = Replace(Mid(Get_ImgData(Mid(strLine, 6)), 2), ";", ",")
            End If
    
    
            If InStr(strLine, "VAL") > 0 Then
                'DIF1
                strDIF1 = strDIF1 & "," & Val(Mid(strLine, 6))
            End If
            If InStr(strLine, "DIF1") > 0 And Len(strLine) >= 100 Then
                Set frmGraph_Tmp = New frmGraph
                strImage = strImage & "^WBC DF1;100;1;" & frmGraph_Tmp.Draw_HMX_DF1(Mid(strDIF1, 2), Mid(strLine, 6))
                Unload frmGraph_Tmp
                Set frmGraph_Tmp = Nothing
            End If
            '====================================================================
        Next
    
        strResult = strResult & "||" & strData & IIf(strImage <> "", vbCrLf & Mid(strImage, 2), "")
        
        lng_S = InStr(mstrCode, strStartChr)
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, strEndChr)
            If lng_E > 0 Then
                lng_E = InStr(lng_E, mstrCode, "")
            End If
        End If
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", _
       Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String

    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function


Private Function Get_PicData(str_in As String, ByVal str_S As String, ByVal str_E As String, ByVal str_Name As String) As String
        '取图形数据
        Dim lng_S As Long, lng_E As Long, int_Len As Integer, i As Integer, Y As Integer
        Dim str_Return As String, varName As Variant, str_Data As String, varline As Variant, str_Line As String
        lng_S = InStr(str_in, str_S)
        lng_E = InStr(str_in, str_E)
        If lng_S > 0 And lng_E > 0 Then
            str_Data = Mid(str_in, lng_S, lng_E - lng_S)
            str_in = Mid(str_in, lng_E)
            
            
            varline = Split(str_Data, vbCrLf)
            For i = LBound(varline) To UBound(varline)
                str_Line = varline(i)
                varName = Split(str_Name, ",")
                
                For Y = LBound(varName) To UBound(varName)
                    int_Len = Len(varName(Y))
                    If Mid(str_Line, 1, int_Len) = varName(Y) Then
                        str_Return = Mid(str_Line, int_Len + 1)
                        str_Return = Mid(str_Return, 2)
                        Do While InStr(str_Return, "") > 0
                            lng_S = InStr(str_Return, "")
                            If lng_S > 0 Then
                                str_Return = Mid(str_Return, 1, lng_S - 5) & Mid(str_Return, lng_S + 4)
                            End If
                        Loop
                        Get_PicData = Get_PicData & "," & str_Return
                    End If
                Next
            Next
        End If
        If Get_PicData <> "" Then Get_PicData = Mid(Get_PicData, 2)
End Function


Private Function Get_ImgData(ByVal str_in As String) As String
    '转换图形数据
    Dim str_Line As String, str_Item As String, str_R As String
    str_Line = str_in
    
    Do While Len(str_Line) > 0
        
        str_Item = Trim(Left(str_Line, 2))
        str_Line = Mid(str_Line, 3)
        
        If str_Item <> "" Then
            str_R = str_R & ";" & CByte("&H" & str_Item)
        End If
    Loop
    If str_R <> "" Then Get_ImgData = str_R
End Function

Private Function DelCRC(ByVal str_in As String) As String
    '取掉CRC校验字符
    Dim str_Line As String
    Dim lng_S As Long
    
    str_Line = str_in
    
    If InStr(str_Line, "") > 0 Then
        Do While InStr(str_Line, "") > 0
            lng_S = InStr(str_Line, "")
            str_Line = Mid(str_Line, 1, lng_S - 5) & Mid(str_Line, lng_S + 4)
'            DelCRC = DelCRC(str_Line)
        Loop
        DelCRC = str_Line
    Else
        DelCRC = str_Line
    End If

    
End Function


