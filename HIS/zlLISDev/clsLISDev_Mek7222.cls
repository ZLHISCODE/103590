VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Mek7222_Bin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:血球仪 Mek 7222
'厂商:日本光电
'作者:陈东
'用户:南京陈玉强
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
    
    Dim varline As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String
    Dim blnEnd As Boolean '是否已解析完
    Dim strImage As String  '图像数据
    Dim varImage As Variant '图像数组
    Dim str散点数据 As String, str散1 As String, str散2 As String, str散3 As String
    Dim lng_S As Long, lng_E As Long
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, ",02")
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, ",03")
    Else
        lng_E = InStr(mstrCode, ",03")
    End If

    If lng_E > 0 And lng_S > 0 Then
        lng_E = lng_E + 4608 * 3 '直方图数据
        lng_E = lng_E + 40000 * 3 '散点图数据
        lng_E = lng_E + 112 * 3 '散点图解析数据
        If Len(Mid(mstrCode, lng_S)) < (lng_E - lng_S) + 3 Then lng_E = 0
    End If
    
    Do While lng_S > 0 And lng_E > 0
    
        lng_E = InStr(lng_S, mstrCode, ",03")

        '取一个完整的记录
        strLine = Mid(mstrCode, lng_S, lng_E - lng_S + 3)
        strImage = Mid(mstrCode, lng_E + 3, 4608 * 3)
        
        str散点数据 = Mid(mstrCode, lng_E + 4608 * 3 + 3, 40112 * 3)
  
        lng_E = lng_E + 4608 * 3 '直方图数据
        lng_E = lng_E + 40000 * 3 '散点图数据
        lng_E = lng_E + 112 * 3 '散点图解析数据
        
        mstrCode = Mid(mstrCode, lng_E + 3)
        
        If Len(strImage) = 4608 * 3 And Len(str散点数据) = 40112 * 3 Then
            WriteLog "收到完整数据", strLine & vbNewLine & strImage & vbNewLine & "散点：" & str散点数据, mstrCode
            strLine = Mid_bin(strLine, 1)
            strImage = Mid_bin(strImage, 1)
            
            varline = Split(strLine, Chr(&HD))
            If strImage <> "" Then
                varImage = Split(strImage, Chr(&HD))
                If UBound(varImage) < 768 Then
                    strLine = ""
                    lng_S = 0: lng_E = 0
                End If
            End If
        Else
            lng_S = 0: lng_E = 0
        End If

        If lng_S > 0 And lng_E > 0 Then
            str日期 = "": str标本号 = "": str检验人标本质控 = "": strItem = "": strTmp = ""
            '解析
            
            '处理结果
            For iRow = 0 To UBound(varline)
                Select Case iRow
                Case 0 '日期
                    If varline(iRow) Like "*##-*-##*" Then
                        str日期 = Mid(varline(iRow), InStr(varline(iRow), "-") - 2, 2)
                        strTmp = Mid(varline(iRow), InStr(varline(iRow), "-") + 1)
                        strTmp = Mid(strTmp, InStr(strTmp, "-") + 1, 2) & "-" & TrunMonth(Mid(strTmp, 1, InStr(strTmp, "-") - 1))
                        str日期 = Format(CDate(strTmp & "-" & str日期), "yyyy-mm-dd")
                    ElseIf varline(iRow) Like "*##/???/##*" Then
                        str日期 = Mid(varline(iRow), InStr(varline(iRow), "/") - 2, 2)
                        strTmp = Mid(varline(iRow), InStr(varline(iRow), "/") + 1)
                        strTmp = Mid(strTmp, InStr(strTmp, "/") + 1, 2) & "-" & TrunMonth(Mid(strTmp, 1, InStr(strTmp, "/") - 1))
                        str日期 = Format(CDate(strTmp & "-" & str日期), "yyyy-mm-dd")
                    Else
                        str日期 = Format(Now, "yyyy-MM-dd")
                    End If
                Case 1 '标本号
                    str标本号 = Val(varline(iRow))
                    str检验人标本质控 = " |血液|0"
                Case 2 'WBC  白细胞总数
                    strItem = "|WBC|" & Format_6318(varline(iRow))
                Case 3 'LYM% 淋巴细胞百分比
                    strItem = strItem & "|LY%|" & Format_6318(varline(iRow))
                Case 4 'MON% 中值细胞百分比
                    strItem = strItem & "|MO%|" & Format_6318(varline(iRow))
                Case 5 'GRA% 中性细胞百分比
                    strItem = strItem & "|NE%|" & Format_6318(varline(iRow))
                Case 6 'LYM# 淋巴细胞绝对值
                    strItem = strItem & "|EO%|" & Format_6318(varline(iRow))
                Case 7 'MON# 中值细胞绝对值
                    strItem = strItem & "|BA%|" & Format_6318(varline(iRow))
                Case 8 'GRA# 中性细胞绝对值
                    strItem = strItem & "|LY|" & Format_6318(varline(iRow))
                Case 9 'EOS# 嗜酸性粒细胞绝对值
                    strItem = strItem & "|MO|" & Format_6318(varline(iRow))
                Case 10 'RBC  红细胞总数
                    strItem = strItem & "|NE|" & Format_6318(varline(iRow))
                Case 11 'HGB 血红蛋白
                    strItem = strItem & "|EO|" & Format_6318(varline(iRow))
                Case 12 'HCT  红细胞压积
                    strItem = strItem & "|BA|" & Format_6318(varline(iRow))
                Case 13 'MCV  平均红细胞体积
                    strItem = strItem & "|RBC|" & Format_6318(varline(iRow))
                Case 14 'MCH  平均血红蛋白含量
                    strItem = strItem & "|HGB|" & Format_6318(varline(iRow))
                Case 15 'MCHC 平均血红蛋白浓度
                    strItem = strItem & "|HCT|" & Format_6318(varline(iRow))
                Case 16 'RDW  红细胞平均宽度
                    strItem = strItem & "|MCV|" & Format_6318(varline(iRow))
                Case 17 'PLT 血小板总数
                    strItem = strItem & "|MCH|" & Format_6318(varline(iRow))
                Case 18 'PCT  血小板压积
                    strItem = strItem & "|MCHC|" & Format_6318(varline(iRow))
                Case 19 'MPV  平均血小板体积
                    strItem = strItem & "|RDW|" & Format_6318(varline(iRow))
                Case 20 'PDW  血小板平均宽度
                    strItem = strItem & "|PLT|" & Format_6318(varline(iRow))
                Case 21 'PDW  血小板平均宽度
                    strItem = strItem & "|PCT|" & Format_6318(varline(iRow))
                Case 22 'PDW  血小板平均宽度
                    strItem = strItem & "|MPV|" & Format_6318(varline(iRow))
                Case 23 'PDW  血小板平均宽度
                    strItem = strItem & "|PDW|" & Format_6318(varline(iRow))
                End Select
            Next
            
            '处理图形
            If strImage <> "" Then
                For iRow = 0 To UBound(varImage)
                    If iRow >= 0 And iRow <= 256 Then
                        If iRow = 0 Then strImage = "^WBC;0;0,271,15-50,51-100,87-,123-200,159-,195-300,235-,271-400"
                        strImage = strImage & ";" & Val(varImage(iRow))
                    End If
                    If iRow > 256 And iRow <= 512 Then
                        If iRow = 257 Then strImage = strImage & "^RBC;0;0,256,18-,56-100,94-,132-200,170-,208-300,246-"
                        strImage = strImage & ";" & Val(varImage(iRow))
                    End If
                    If iRow > 512 And iRow <= 768 Then
                        If iRow = 513 Then strImage = strImage & "^PLT;0;0,220,9-2,42-,75-20,108-,141-40,174-,207-60"
                        strImage = strImage & ";" & Val(varImage(iRow))
                    End If
                Next
            End If
            
            If strItem <> "" And Val(str标本号) <> 0 Then
                strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
                
                '处理散点图
                If str散点数据 <> "" Then
                    str散1 = SaveBmp(str日期, str标本号, str散点数据)
                    If str散1 <> "" Then
                        strImage = strImage & str散1
                    End If
                End If
                
                If strImage <> "" Then
                    strResult = strResult & vbCrLf & Mid(strImage, 2)
                End If
            End If
    
            lng_S = InStr(mstrCode, ",02")
            If lng_S > 0 Then
                lng_E = InStr(lng_S, mstrCode, ",03")
            Else
                lng_E = InStr(mstrCode, ",03")
            End If
        
            If lng_E > 0 And lng_S > 0 Then
                lng_E = lng_E + 4608 * 3 '直方图数据
                lng_E = lng_E + 40000 * 3 '散点图数据
                lng_E = lng_E + 112 * 3 '散点图解析数据
                If Len(Mid(mstrCode, lng_S)) < (lng_E - lng_S) + 3 Then lng_E = 0
            End If
        End If
    Loop


    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If

End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:解析需要向仪器发送的标本信息
    '           标本之间以||分隔
    '           元素之间以|分隔
    '           第0个元素：检验时间
    '           第1个元素：样本序号
    '           第2个元素：检验人
    '           第3个元素：标本
    '           第4个元素：是否质控品
    '           第5～9元素：系统保留
    '           从第10个元素开始为需要的检验项目。
    '   iSendStep  : 步骤
    '   blnSuccess : 是否传输失败
    '   strResponse: 仪器的应答信息
    '   blnUndo    : 是否取消.
    '
    '   返回值：发到仪器的字串或指令。
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function


Private Function TrunMonth(ByVal strMonth As String) As String
    '转换英文的月份为数字
    Select Case Mid(UCase(strMonth), 1, 3)
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select

End Function

Private Function Format_6318(ByVal strVal As String) As String
    '格式转换
    '有>,<号原样反回，否则转为数字返回
    If InStr(strVal, ">") > 0 Or InStr(strVal, "<") > 0 Then
        Format_6318 = Trim(strVal)
    Else
        Format_6318 = Val(Trim(strVal))
    End If
End Function



Private Function SaveBmp(ByVal str日期 As String, ByVal str标本号 As String, ByVal strData As String) As String
    '生成3个散点图的BMP文件
    Dim frmPic1 As New frmGraph, strFile1 As String
    Dim frmPic2 As New frmGraph, strFile2 As String
    Dim frmPic3 As New frmGraph, strFile3 As String
    
    Dim str散1 As String, str散2 As String, str散3 As String, strTmp As String
    Dim lngX1 As Long, lngX2 As Long, lngY1 As Long, lngY2 As Long
    Dim lngloop As Long
    Dim strIn As String
    
    strIn = strData
    If Len(strIn) <> 40112 * 3 Then Exit Function
    
    frmPic1.Picture1.Scale (0, 0)-(256, 256)
    frmPic1.Picture1.BackColor = vbWhite
    
    frmPic2.Picture1.Scale (0, 0)-(256, 256)
    frmPic2.Picture1.BackColor = vbWhite
    
    frmPic3.Picture1.Scale (0, 0)-(256, 256)
    frmPic3.Picture1.BackColor = vbWhite
    
    For lngloop = 1 To 10000
        strTmp = Mid(strIn, 1, 12)
        str散1 = CByte("&H" & Mid(strTmp, 2, 2)) 'Size表示血细胞的大小
        str散2 = CByte("&H" & Mid(strTmp, 5, 2)) 'Complexity表示血细胞的复杂度
        str散3 = CByte("&H" & Mid(strTmp, 8, 2)) 'Granularity表示血细胞的颗粒的程度
        strIn = Mid(strIn, 13)
        '纵轴是Size、横轴是Complexity的散点图（以下称为S-C散点图）；
        frmPic1.Picture1.PSet (Val(str散2), 256 - Val(str散1)), &H80000011  's-c图
        '纵轴是Size、横轴是Granularity的嗜中性白细胞/嗜酸性白细胞分类用散点图（以下称为NE-EO分类用散点图）
        If Val(str散3) < 35 And Val(str散2) < 48 Then
            frmPic3.Picture1.PSet (Val(str散3), 256 - Val(str散1)), &H80000011  'NE-EO图
        End If
        If Val(str散3) >= 36 Or Val(str散2) >= 50 Then
        '纵轴是Size、横轴是Granularity的单核细胞/嗜碱性白细胞分类用散点图（以下称为MO-BA分类用散点图）；
            frmPic2.Picture1.PSet (Val(str散3), 256 - Val(str散1)), &H80000011  'MO-BA图
        End If
    Next
    str散1 = ""
    strFile1 = App.Path & "\" & Replace(str日期, "-", "") & "_" & str标本号 & "_S1.JPG"
    strFile2 = App.Path & "\" & Replace(str日期, "-", "") & "_" & str标本号 & "_S2.JPG"
    strFile3 = App.Path & "\" & Replace(str日期, "-", "") & "_" & str标本号 & "_S3.JPG"
    
    'pic1的坐标
    lngX1 = 255: lngY1 = 44
    lngX2 = 128: lngY2 = 44
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 128: lngY1 = 44
    lngX2 = 65: lngY2 = 90
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 65: lngY1 = 90
    lngX2 = 50: lngY2 = 90
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 20: lngY1 = 255
    lngX2 = 20: lngY2 = 70
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 21: lngY1 = 70
    lngX2 = 30: lngY2 = 58
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 30: lngY1 = 58
    lngX2 = 35: lngY2 = 56
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 36: lngY1 = 56
    lngX2 = 50: lngY2 = 71
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 50: lngY1 = 71
    lngX2 = 50: lngY2 = 81
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 52: lngY1 = 81
    lngX2 = 52: lngY2 = 102
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 54: lngY1 = 102
    lngX2 = 54: lngY2 = 152
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 54: lngY1 = 152
    lngX2 = 90: lngY2 = 255
    frmPic1.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)
    
    'pic2的坐标
    lngX1 = 255: lngY1 = 44
    lngX2 = 0: lngY2 = 44
    frmPic2.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 0: lngY1 = 132
    lngX2 = 82: lngY2 = 132
    frmPic2.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 82: lngY1 = 132
    lngX2 = 177: lngY2 = 44
    frmPic2.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)

    lngX1 = 82: lngY1 = 132
    lngX2 = 250: lngY2 = 255
    frmPic2.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)
    
    'pic3的坐标
    lngX1 = 255: lngY1 = 60
    lngX2 = 0: lngY2 = 60
    frmPic3.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)
    
    lngX1 = 255: lngY1 = 150
    lngX2 = 0: lngY2 = 150
    frmPic3.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)
    
    lngX1 = 50: lngY1 = 255
    lngX2 = 50: lngY2 = 60
    frmPic3.Picture1.Line (lngX1, 256 - lngY1)-(lngX2, 256 - lngY2)
'
'    SavePicture frmPic1.Picture1.Image,  APP.Path "\TMP1.BMP"
'    SavePicture frmPic2.Picture1.Image, strFile2
'    SavePicture frmPic3.Picture1.Image, strFile3
    
    SavePic frmPic1.Picture1.Image, strFile1, "JPG"
    SavePic frmPic2.Picture1.Image, strFile2, "JPG"
    SavePic frmPic3.Picture1.Image, strFile3, "JPG"
    
    SaveBmp = "^Wbc1;123;1;" & strFile1 & "^Wbc2;123;1;" & strFile2 & "^Wbc3;123;1;" & strFile3
    Exit Function

ErrHand:
    SaveBmp = ""
End Function
