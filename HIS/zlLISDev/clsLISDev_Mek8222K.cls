VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Mek8222K"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Mek 822k
'厂商:日本光电
'作者:陈玉强
'用户:栖霞医院rs-232传输数据
'-- 返回图形数据时的格式:
    '补充图像的方式：
    '                   1.图像数据跟随指标数据后，使用回车换行符来分隔。
    '                   2.有多个图像数据时使用"^"来分隔
    '                   3.单个图像数据格式: 图像画法 0=直方图  1=散点图  2=血流变粘度特征曲线  3=血沉曲线  100=图片数据
    '                     0) 直方图: 图像名称;图像画法(0=直方图  1=散点图);Y1;Y2;Y3;Y4;Y5...
    '                     1) 散点图: 图像名称;图像画法(0=直方图  1=散点图):
    '                        例:00000100001000010000100010;00000100001000010000100010;
    '                        说明:1.散点图以点阵方式保存每一行使用分号来分隔.
    '                             2.有多少个分号就有多少行
    '                             3.每一行有多少个点由每一行的长度来确定
    '                             3.画图的方向是从最上边向下画，如有65*65的图就是从65行开始画(最上边开始画)
    '                     2) 粘度特征曲线:图像名称;图像画法;座标数据;曲线及描点数据;坐标轴标题数据
    '                                   其中  座标数据：Y长度,X长度|X座标-X座标显示的数字,....|Y座标-Y座标显示的数字,....
    '                                   曲线及描点数据:粘度曲线1的高点和低点座标|粘度曲线2的高点和低点座标~低切点坐标,中切点坐标,高切点坐标
    '                                   坐标轴标题数据:Y坐标标题文字,X坐标,Y座标~X坐标标题文字,X坐标,Y座标
    '                        例:粘度特征曲线;2;20,200|20-20,40-40,60-60,80-80,100-100,120-120,140-140,160-160,180-180,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20;9.25,10,4.4,150|6.5,10,3.65,150~10-8.989,60-4.803,150-4.05;VIS(mPa.s),25,20~SHR(1/S),195,1
    '                     3) 血沉曲线:图像名称;图像画法;座标数据;描点数据;坐标轴标题数据
    '                                   其中  座标数据：Y长度,X长度|X座标-X座标显示的数字,....|Y座标-Y座标显示的数字,....
    '                                   描点数据:血沉值1,血沉值2,....血沉值30
    '                                   坐标轴标题数据:Y坐标标题文字,X坐标,Y座标~X坐标标题文字,X坐标,Y座标
    '                        例:血沉曲线;3;36,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36;.5,.5,1,1,1,1.5,1.5,2,2,2,2.5,3,3,3.5,4,4.5,5.5,6.5,8,9,10.5,11.5,12.5,13.5,14.5,15.5,16.5,18,19,20;血沉值(mm),5,36~时间(m),55,1
    '                     4) PLT图：图像名称;图像画法;座标数据;描点数据
    '                               其中 坐标数据：Y长度,X长度,X座标-X座标显示的数字,....
    '                                    描点数据: Y1,Y2,Y3,......|Y1,Y2,Y3,......
    '                        例:PLT;4;200,262;0,0,0,0,0,0,0,0,0,0,0,0,0,0,3,3,4,4,7,7,12,12,17,17,20,20,25,25,30,30,33,33,36,36,41,41,43,43,44,44,46,46,47,47,47,47,47,47,46,46,46,46,44,44,44,44,43,43,41,41,39,39,38,38,36,36,35,35,33,33,31,31,30,30,28,28,27,27,25,25,23,23,22,22,22,22,20,20,19,19,17,17,15,15,15,15,14,14,12,12,12,12,11,11,11,11,9,9,9,9,9,9,7,7,7,7,7,7,6,6,6,6,6,6,4,4,4,4,4,4,4,4,3,3,3,3,3,3,3,3,3,3,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,7,7,9,9,8,8,9,9,12,12,16,16,22,22,26,26,30,30,35,35,36,36,37,37,39,39,42,42,44,44,46,46,46,46,44,44,43,43,40,40,37,37,37,37,37,37,39,39,37,37,36,36,32,32,29,29,25,25,23,23,22,22,22,22,21,21,19,19,18,18,16,16,16,16,15,15,15,15,15,15,14,14,12,12,11,11,9,9,9,9,8,8,8,8,7,7,7,7,7,7,7,7,7,7,8,8,7,7,7,7,5,5,4,4,4,4,2,2,4,4,4,4,2,2,2,2,4,4
    '                   100) 图片数据:图像名称;图像画法;[读取数据后是否删除];全路径
    '                        例:WBC Fsc;100;1;C:\tempfile.gif
    


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
    On Error GoTo errHand '打开错误陷阱
    '------------------------------------------------
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

    Dim varLine As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String
    Dim blnEnd As Boolean '是否已解析完
    Dim strImage As String  '图像数据
    Dim varImage As Variant '图像数组
    Dim intItem As Integer
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    If strSource Like "*MEK-8222**" Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            '取一个完整的记录,剩余的存到strReserved中,下次解析
            strLine = Mid(strReserved, InStr(strReserved, "MEK-8222"), InStr(strReserved, ""))
            strReserved = Mid(strReserved, InStr(strReserved, "") + 1)
            str日期 = "": str标本号 = "": str检验人标本质控 = "": strItem = "": strTmp = "": intItem = 0
            '解析
            varLine = Split(strLine, Chr(13))

            '处理结果
            For iRow = 0 To UBound(varLine)
                Select Case iRow
                    Case 15, 16, 17 '日期
                    If iRow = 15 Then
                        str日期 = varLine(iRow) & "-"
                    ElseIf iRow = 16 Then
                        str日期 = str日期 & varLine(iRow) & "-"
                    Else
                        str日期 = Format(str日期 & varLine(iRow), "yyyy-MM-dd")
                    End If
                    Case 22 '标本号
                    str标本号 = Val(varLine(iRow))
                    str检验人标本质控 = " |血液|0"
                    Case 23 To 44
                    intItem = intItem + 1
                    strTmp = Replace(Replace(Trim(varLine(iRow)), "L", ""), "H", "")

                    strItem = strItem & "|" & intItem & "|" & Val(strTmp)
                End Select
            Next

            If strItem <> "" And Val(str标本号) <> 0 Then
                strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
            End If
            '            If Not strReserved Like "*MEK-8222**" Then
            blnEnd = True
            '            End If
        Loop
    Else
        '收到的信息不全,返回,等下次再解析
        strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If

    '------------------------------------------------
    Exit Sub
    '----------------
errHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLisDev_DSI905"
    Close #lngErrFileNo
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
