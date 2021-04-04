VERSION 5.00
Begin VB.Form frmReadInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "读卡"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4035
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   255
      Top             =   120
   End
End
Attribute VB_Name = "frmReadInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public glngParentHwnd As Long
Private mstrPath As String          '读卡数据文件路径

Private mstrID As String
Private mstrName As String
Private mstrSex As String
Private mstrNation As String
Private mdatBirthDay As Date
Private mstrAddress As String
Private mpicPhoto As StdPicture     '身份证照片信息
Private mblnAutoRead As Boolean  '是否自动读卡——调用 SetEnable 方法为自动读卡，调用 ReadIDCard 方法为 手动读卡

Private Enum ReadMode
    Base = 1        '形成文字信息文件WZ.TXT、相片文件XP.WLT和ZP.BMP
    onlytext = 2    '形成文字信息文件WZ.TXT和相片文件XP.WLT
    NewAdd = 3      '形成最新住址文件NEWADD.TXT
End Enum

Private Const TXTFile = "\wz.txt"
Private Const BMPFile = "\zp.bmp"
Private Const WLTFile = "\zp.wlt"
Private pucManaMsg As String * 4
Private Const IfOpen = 0 '0表示不在该函数内部打开和关闭串口，此时确保之前调用了Syn_OpenPort来打开端口，并且在不需要与端口通信时，调用Syn_ClosePort关闭端口；
                        '非0表示在API函数内部包含了打开端口和关闭端口函数，之前不需要调用Syn_OpenPort，也不用再调用Syn_ClosePort?
                        
'Private pucIIN As Integer, pucSN As Integer, puiCHMsgLen As Integer, puiPHMsgLen As Integer, iIfOpen As Integer
Private pucIIN As String * 8
Private pucSN As String * 8
Private puiCHMsgLen As Long
Private puiPHMsgLen As Long
Private iIfOpen As Integer

Private lngReturn As Integer '定义返回结果值
Private mblnCancel As Boolean
Public mobjIDCard As clsIDCard

Private Const GWL_STYLE = (-16)
Private Const WS_DISABLED = &H8000000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long


Private Sub Form_Load()
    iIfOpen = 1
    mblnAutoRead = True         '程序默认为自动读卡
    mstrPath = GetSetting("ZLSOFT", "公共全局", "程序路径", "C:")
    If mstrPath <> "C:" Then mstrPath = Mid(mstrPath, 1, InStrRev(mstrPath, "\") - 1)

    If Dir(mstrPath, vbDirectory) = "" Then
        MsgBox "ZLHIS应用程序目录和C盘都不存在!不能读卡", vbInformation, App.ProductName
        Timer1.Enabled = False
    End If
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    If (GetWindowLong(mobjIDCard.GetParent, GWL_STYLE) And WS_DISABLED) <> WS_DISABLED Then
        '107213:李南春,2017/4/12,GDI增长问题,为避免卡死，第二次轮询才开始读卡
        If mobjIDCard.GetParent <> 0 Then
            If GetActiveWindow <> glngParentHwnd Then mblnCancel = True: Exit Sub
            If mblnCancel Then mblnCancel = False: Exit Sub '跳过一次
        End If
        Select Case glngType
            Case IDCardType.CVR100U, IDCardType.CVR100D, IDCardType.SS_V1, IDCardType.XZX_KDQ, IDCardType.GTICR100, IDCardType.GTICR100_01, _
                IDCardType.GTICR100_1
                If Authenticate = 1 Then Call ReadIDCard
            Case IDCardType.CVR100U_1, IDCardType.CVR100D_1
                If CVR_Authenticate = 1 Then Call ReadIDCard
            Case IDCardType.HX_FDX9
                If SDT_StartFindIDCard(1, "", 1) = CByte(&H9F) Then Call ReadIDCard
            Case IDCardType.DKQ_116D
                lngReturn = Syn_ClosePort(1001)
                lngReturn = Syn_OpenPort(1001)
                If lngReturn = 0 Then Call ReadIDCard
            Case IDCardType.CVR100
                lngReturn = SDT_StartFindIDCard(1001, "", 1)
                If lngReturn = CByte(&H9F) Then Call ReadIDCard
            Case IDCardType.COMMON
                '找卡
                i = SDT_StartFindIDCard(editPort, pucIIN, iIfOpen)
                If i <> CByte(&H9F) Then
                    '再找卡
                    i = SDT_StartFindIDCard(editPort, "", 1)
                    If i <> CByte(&H9F) Then
                        i = SDT_ClosePort(editPort)
                    Else
                        Call ReadIDCard
                    End If
                Else
                    Call ReadIDCard
                End If
            Case IDCardType.SS728M01_B01C
                Call ReadIDCard
        End Select
    End If
End Sub


Private Sub ReadIDCard()
    Dim intTmp As Integer, strMSG As String
    Dim strPucManaMsg As String
    Dim i As Integer
    
    Select Case glngType
        Case IDCardType.CVR100U, IDCardType.CVR100D, IDCardType.SS_V1, IDCardType.XZX_KDQ, IDCardType.GTICR100
            intTmp = Read_Content_Path(mstrPath, ReadMode.Base)
            Select Case intTmp
                Case 0
                    strMSG = "请将身份证停留在设备可识别位置,至少保持1秒!"
                Case 1
                    strMSG = "" '成功
                Case 2
                    strMSG = "没有最新住址信息!"
                Case -1
                    strMSG = "相片解码错误!"
                Case -2
                    strMSG = "wlt文件后缀错误!"
                Case -3
                    strMSG = "wlt文件打开错误!"
                Case -4
                    strMSG = "wlt文件格式错误!"
                Case -5
                    strMSG = "软件未授权!"
                Case -11
                    strMSG = "无效参数!"
                Case -12
                    strMSG = "路径太长!"
                Case Else
                    strMSG = "设备未知错误!"
            End Select
        Case IDCardType.GTICR100_1, IDCardType.GTICR100_01
            intTmp = Read_Content(ReadMode.Base)
            Select Case intTmp
                Case 0
                    strMSG = "请将身份证停留在设备可识别位置,至少保持1秒!"
                Case 1
                    strMSG = "" '成功
                Case 2
                    strMSG = "没有最新住址信息!"
                Case -1
                    strMSG = "相片解码错误!"
                Case -2
                    strMSG = "wlt文件后缀错误!"
                Case -3
                    strMSG = "wlt文件打开错误!"
                Case -4
                    strMSG = "wlt文件格式错误!"
                Case -5
                    strMSG = "软件未授权!"
                Case -11
                    strMSG = "无效参数!"
                Case -12
                    strMSG = "路径太长!"
                Case Else
                    strMSG = "设备未知错误!"
            End Select
        Case IDCardType.CVR100U_1, IDCardType.CVR100D_1
            intTmp = CVR_Read_Content(ReadMode.Base)
            Select Case intTmp
                Case 0
                    strMSG = "请将身份证停留在设备可识别位置,至少保持1秒!"
                Case 1
                    strMSG = "" '成功
                Case 2
                    strMSG = "没有最新住址信息!"
                Case -1
                    strMSG = "相片解码错误!"
                Case -2
                    strMSG = "wlt文件后缀错误!"
                Case -3
                    strMSG = "wlt文件打开错误!"
                Case -4
                    strMSG = "wlt文件格式错误!"
                Case -5
                    strMSG = "软件未授权!"
                Case -11
                    strMSG = "无效参数!"
                Case -12
                    strMSG = "路径太长!"
                Case Else
                    strMSG = "设备未知错误!"
            End Select
        Case IDCardType.HX_FDX9
            intTmp = SDT_SelectIDCard(1, strPucManaMsg, 1)
            Select Case intTmp
                Case CByte(&H90)
                    strMSG = "" '成功
                    intTmp = SDT_ReadBaseMsgToFile(1, mstrPath & TXTFile, puiCHMsgLen, mstrPath & WLTFile, puiPHMsgLen, 1)
                    If intTmp = CByte(&H90) Then
                        intTmp = GetBmp(mstrPath & WLTFile, 1)
                        If intTmp <> 1 Then
                            Timer1.Enabled = False
                            strMSG = "照片解析失败！"
                            MsgBox strMSG, vbInformation, App.ProductName
                            Timer1.Enabled = True
                        End If
                    Else
                        Timer1.Enabled = False
                        strMSG = "读卡失败！"
                        MsgBox strMSG, vbInformation, App.ProductName
                        Timer1.Enabled = True
                    End If
                Case CByte(&H81)
                    strMSG = "选卡失败！"
            End Select
        Case IDCardType.CVR100
            intTmp = SDT_SelectIDCard(1001, strPucManaMsg, 1)
            Select Case intTmp
                Case CByte(&H90)
                    strMSG = "" '成功
                    intTmp = SDT_ReadBaseMsgToFile(1001, mstrPath & TXTFile, puiCHMsgLen, mstrPath & WLTFile, puiPHMsgLen, 1)
                    If intTmp = CByte(&H90) Then
                        intTmp = GetBmp(mstrPath & WLTFile, 1)
                        If intTmp <> 1 Then
                            Timer1.Enabled = False
                            strMSG = "照片解析失败！"
                            MsgBox strMSG, vbInformation, App.ProductName
                            Timer1.Enabled = True
                        End If
                    Else
                        Timer1.Enabled = False
                        strMSG = "读卡失败！"
                        MsgBox strMSG, vbInformation, App.ProductName
                        Timer1.Enabled = True
                    End If
                Case CByte(&H81)
                    strMSG = "选卡失败！"
            End Select
        Case IDCardType.DKQ_116D
            Call ClearIDCardInfor   '临时清空结构体，由于Syn_ReadMsg函数未返回0时，仍然可以读取非照片字段
            intTmp = Syn_StartFindIDCard(1001, pucManaMsg, IfOpen)
            intTmp = Syn_SelectIDCard(1001, pucManaMsg, IfOpen)
            intTmp = Syn_ReadMsg(1001, IfOpen, IDCardInfor)
            Call Syn_ClosePort(1001)   '关闭端口，防止多窗口调用时冲突
            Select Case intTmp
                Case 0
                    strMSG = "" '操作成功或相片解码解码正确
                Case -1
                    strMSG = "端口打开失败/端口尚未打开/端口号不合法"
                Case -2
                    strMSG = "证/卡中此项无内容"
                Case -3
                    strMSG = "PC接收超时，在规定的时间内未接收到规定长度的数据"
                Case -4
                    strMSG = "数据传输错误"
                Case -5
                    strMSG = "该SAM_V串口不可用，只在SDT_GetCOMBaud时才有可能返回"
                Case -6
                    strMSG = "接收业务终端数据的校验和错"
                Case -7
                    strMSG = "接收业务终端数据的长度错"
                Case -8
                    strMSG = "接收业务终端的命令错误，包括命令中的各种数值或逻辑搭配错误"
                Case -9
                    strMSG = "越权操作"
                Case -10
                    strMSG = "无法识别的错误"
                Case -11
                    strMSG = "寻找证/卡失败"
                Case -12
                    strMSG = "选取证/卡失败"
                Case -13
                    strMSG = "调用sdtapi.dll错误"
                Case -14
                    strMSG = "相片解码错误"
                Case -15
                    strMSG = "授权文件不存在"
                Case -16
                    strMSG = "设备连接错误"
            End Select
            If TrimStr(IDCardInfor.Name) <> "" Then strMSG = ""  '有时读卡返回失败，但除照片信息外仍然可以读取到
        Case IDCardType.COMMON
            '选卡
            i = SDT_SelectIDCard(editPort, pucSN, iIfOpen)
            If i <> CByte(&H90) Then
                strMSG = "选卡失败，请重新放卡"
                Call SDT_ClosePort(editPort)
            Else
                '读卡
                intTmp = SDT_ReadBaseMsgToFile(editPort, mstrPath & TXTFile, puiCHMsgLen, mstrPath & WLTFile, puiPHMsgLen, iIfOpen)
                If intTmp = CByte(&H90) Then
                    intTmp = GetBmp(mstrPath & WLTFile, 1)
                    If intTmp <> 1 Then
                        Timer1.Enabled = False
                        Call SDT_ClosePort(editPort)
                        Select Case intTmp
                            Case 0
                                strMSG = "调用sdtapi.dll错误"
                            Case 1
                                '正常
                            Case -1
                                strMSG = "相片解码错误！"
                            Case -2
                                strMSG = "wlt文件后缀错误！"
                            Case -3
                                strMSG = "wlt文件打开错误！"
                            Case -4
                                strMSG = "wlt文件格式错误！"
                            Case -5
                                strMSG = "软件未授权！"
                            Case -6
                                strMSG = "设备连接错误！"
                        End Select
                        MsgBox strMSG, vbInformation, App.ProductName
                        Timer1.Enabled = True
                    End If
                Else
                    Timer1.Enabled = False
                    Call SDT_ClosePort(editPort)
                    strMSG = "读卡失败！"
                    MsgBox strMSG, vbInformation, App.ProductName
                    Timer1.Enabled = True
                End If
            End If
        Case IDCardType.SS728M01_B01C
            If Not ReadSS728M01 Then strMSG = "读取身份证失败！"
    End Select
    If strMSG = "" Then
        Call ReadInfoFromFile
        Call DelIDCardFile
    End If

End Sub

Private Sub ReadInfoFromFile()
    Dim strID As String, strName As String, strSex As String
    Dim strNation As String, datBirthday As Date, strAddress As String

    Dim tmp1 As Byte, tmp2 As Byte, intTmp As Integer
    Dim strData As String, strBirthDay As String

    mstrID = "": mstrName = "": mstrSex = "": mstrNation = "":  mdatBirthDay = datBirthday: mstrAddress = ""
    Set mpicPhoto = Nothing

    Select Case glngType
        Case IDCardType.DKQ_116D
            '身份号码
            strID = TrimStr(IDCardInfor.IDcardno)
            '姓名
            strName = TrimStr(IDCardInfor.Name)
            '性别
            strSex = TrimStr(IDCardInfor.sex)
           '代码转换
            Select Case strSex
                Case "1"
                    strSex = "男"
                Case "2"
                    strSex = "女"
                Case Else
                    strSex = "未知"
            End Select
            '民族
            strNation = TrimStr(IDCardInfor.nation)
            strNation = TranNation(Val(strNation))
            '出生日期
            datBirthday = CDate(Mid(IDCardInfor.born, 1, 4) & "-" & Mid(IDCardInfor.born, 5, 2) & "-" & Mid(IDCardInfor.born, 7, 2)) 'Format(TrimStr(IDCardInfor.born), "yyyy-MM-dd")
            '住址
            strAddress = TrimStr(IDCardInfor.address)
            Set mpicPhoto = LoadPicture(TrimStr(IDCardInfor.PhotoFileName))
        Case IDCardType.SS728M01_B01C
            '身份号码
            strID = SS728M01.ss_id_query_number
            '姓名
            strName = SS728M01.ss_id_query_name
            '性别
            strSex = SS728M01.ss_id_query_sex
            '民族
            strNation = SS728M01.ss_id_query_folk
            '出生日期
            If Len(SS728M01.ss_id_query_birth) >= 8 Then datBirthday = CDate(Mid(SS728M01.ss_id_query_birth, 1, 4) & "-" & Mid(SS728M01.ss_id_query_birth, 5, 2) & "-" & Mid(SS728M01.ss_id_query_birth, 7, 2))
            '住址
            strAddress = TrimStr(IIf(SS728M01.ss_id_query_newaddr = "", SS728M01.ss_id_query_address, SS728M01.ss_id_query_newaddr))
            Set mpicPhoto = LoadPicture(TrimStr(SS728M01.ss_id_query_photofile))
        Case Else
            Open IIf(IDCardType.GTICR100_1 = glngType, App.Path & TXTFile, mstrPath & TXTFile) For Binary As #1
                Do While Not EOF(1)   ' 检查文件尾。
                    Get #1, , tmp1
                    Get #1, , tmp2
                    strData = strData & ChrW(tmp2 * CLng(256) + tmp1)
                Loop
            Close #1

            Open IIf(IDCardType.GTICR100_01 = glngType, App.Path & TXTFile, mstrPath & TXTFile) For Binary As #1
                Do While Not EOF(1)   ' 检查文件尾。
                    Get #1, , tmp1
                    Get #1, , tmp2
                    strData = strData & ChrW(tmp2 * CLng(256) + tmp1)
                Loop
            Close #1



            '身份号码
            strID = Trim(Mid(strData, 62, 18))
            '姓名
            strName = Trim(Mid(strData, 1, 15))
            '性别
            strSex = Mid(strData, 16, 1)
            '民族
            strNation = Mid(strData, 17, 2)
            '出生日期
            strBirthDay = Mid(strData, 19, 8)
            '住址
            strAddress = Trim(Mid(strData, 27, 35))


            '代码转换
            Select Case strSex
                Case "0"
                    strSex = "未知"
                Case "1"
                    strSex = "男"
                Case "2"
                    strSex = "女"
                Case Else
                    strSex = "未说明"
            End Select
            strNation = GetNation(strNation)
            If IsNumeric(strBirthDay) And Len(strBirthDay) = 8 Then
                datBirthday = CDate(Mid(strBirthDay, 1, 4) & "-" & Mid(strBirthDay, 5, 2) & "-" & Mid(strBirthDay, 7, 2))
            End If

            Set mpicPhoto = LoadPicture(IIf(IDCardType.GTICR100_1 = glngType, App.Path & BMPFile, mstrPath & BMPFile))
    End Select

    If mblnAutoRead = False Then
        mstrID = strID: mstrName = strName: mstrSex = strSex: mstrNation = strNation: mdatBirthDay = datBirthday: mstrAddress = strAddress
    Else
        Call mobjIDCard.ShowIDCardInfo(strID, strName, strSex, strNation, datBirthday, strAddress)
    End If
'    Set mpicPhoto = Nothing
End Sub

Public Sub DelIDCardFile()
    If Dir(mstrPath & TXTFile) <> "" Then Call Kill(mstrPath & TXTFile)
    If Dir(mstrPath & BMPFile) <> "" Then Call Kill(mstrPath & BMPFile)
    '以下删除文件代码主要针对 GTICR100
    If Dir(App.Path & TXTFile) <> "" Then Call Kill(App.Path & TXTFile)
    If Dir(App.Path & BMPFile) <> "" Then Call Kill(App.Path & BMPFile)
    '针对新中新
    If Dir(TrimStr(IDCardInfor.PhotoFileName)) <> "" And TrimStr(IDCardInfor.PhotoFileName) <> "" Then Call Kill(TrimStr(IDCardInfor.PhotoFileName))
End Sub

'民族代码查表
Public Function GetNation(ByVal strNationcode As String) As String
    Dim strNationArray As Variant

    strNationArray = Array("汉", "蒙古", "回", "藏", "维吾尔", "苗", "彝", "壮", "布依", "朝鲜", _
                        "满", "侗", "瑶", "白", "土家", "哈尼", "哈萨克", "傣", "黎", "傈僳", _
                        "佤", "畲", "高山", "拉祜", "水", "东乡", "纳西", "景颇", "柯尔克孜", "土", _
                        "达斡尔", "仫佬", "羌", "布朗", "撒拉", "毛南", "仡佬", "锡伯", "阿昌", "普米", _
                        "塔吉克", "怒", "乌孜别克", "俄罗斯", "鄂温克", "德昂", "保安", "裕固", "京", "塔塔尔", _
                        "独龙", "鄂伦春", "赫哲", "门巴", "珞巴", "基诺")

    If Trim(strNationcode) <> "" Then
        If ((CByte(Trim(strNationcode)) - 1) >= 0) And ((CByte(Trim(strNationcode)) - 1) <= 55) Then
            GetNation = strNationArray(CByte(Trim(strNationcode)) - 1)
            '90373:李南春，2015/11/6,返回民族全称
            GetNation = GetNation & "族"
        Else
            GetNation = "其他"
        End If
    End If
End Function

Public Sub Read_Card(strID As String, strName As String, strSex As String, _
                             strNation As String, datBirthday As Date, strAddress As String)
    mblnAutoRead = False            '调用此方法时为手动读卡
    mstrID = "": mstrName = "": mstrSex = "": mstrNation = "": mdatBirthDay = datBirthday: mstrAddress = ""
    Set mpicPhoto = Nothing
    Dim i As Integer
    Select Case glngType
        Case IDCardType.CVR100U, IDCardType.CVR100D, IDCardType.SS_V1, IDCardType.XZX_KDQ, IDCardType.GTICR100, IDCardType.GTICR100_01, IDCardType.GTICR100_1
            If Authenticate = 1 Then Call ReadIDCard
        Case IDCardType.CVR100U_1, IDCardType.CVR100D_1
            If CVR_Authenticate = 1 Then Call ReadIDCard
        Case IDCardType.HX_FDX9
            If SDT_StartFindIDCard(1, "", 1) = CByte(&H9F) Then Call ReadIDCard
        Case IDCardType.DKQ_116D
            lngReturn = Syn_ClosePort(1001)
            lngReturn = Syn_OpenPort(1001)
            If lngReturn = 0 Then Call ReadIDCard
        Case IDCardType.CVR100
            lngReturn = SDT_StartFindIDCard(1001, "", 1)
            If lngReturn = CByte(&H9F) Then Call ReadIDCard
        Case IDCardType.COMMON
            '找卡
            i = SDT_StartFindIDCard(editPort, pucIIN, iIfOpen)
            If i <> CByte(&H9F) Then
                '再找卡
                i = SDT_StartFindIDCard(editPort, pucIIN, iIfOpen)
                If i <> CByte(&H9F) Then
                    Call SDT_ClosePort(editPort)
                Else
                    Call ReadIDCard
                End If
            End If
    End Select

    strID = mstrID: strName = mstrName: strSex = mstrSex: strNation = mstrNation: datBirthday = mdatBirthDay: strAddress = mstrAddress

    mblnAutoRead = True             '恢复为自动读卡
End Sub

Public Function ReadPhotoInfo() As StdPicture
    Set ReadPhotoInfo = mpicPhoto
End Function

Public Function TranNation(ByVal lngNo As Long) As String
    Dim strNation As String
    Select Case lngNo
    Case 1
        strNation = "汉族"
    Case 2
        strNation = "蒙古族"
    Case 3
        strNation = "回族"
    Case 4
        strNation = "藏族"
    Case 5
        strNation = "维吾尔族"
    Case 6
        strNation = "苗族"
    Case 7
        strNation = "彝族"
    Case 8
        strNation = "壮族"
    Case 9
        strNation = "布依族"
    Case 10
        strNation = "朝鲜族"
    Case 11
        strNation = "满族"
    Case 12
        strNation = "侗族"
    Case 13
        strNation = "瑶族"
    Case 15
        strNation = "土家族"
    Case 16
        strNation = "哈尼族"
    Case 17
        strNation = "哈萨克族"
    Case 18
        strNation = "傣族"
    Case 19
        strNation = "黎族"
    Case 20
        strNation = "傈僳族"
    Case 21
        strNation = "佤族"
    Case 22
        strNation = "畲族"
    Case 23
        strNation = "高山族"
    Case 24
        strNation = "拉祜族"
    Case 25
        strNation = "水族"
    Case 26
        strNation = "东乡族"
    Case 27
        strNation = "纳西族"
    Case 28
        strNation = "景颇族"
    Case 29
        strNation = "柯尔克孜族"
    Case 30
        strNation = "土族"
    Case 31
        strNation = "达斡尔族"
    Case 32
        strNation = "仫佬族"
    Case 33
        strNation = "羌族"
    Case 34
        strNation = "布朗族"
    Case 35
        strNation = "撒拉族"
    Case 36
        strNation = "毛南族"
    Case 37
        strNation = "仡佬族"
    Case 38
        strNation = "锡伯族"
    Case 39
        strNation = "阿昌族"
    Case 40
        strNation = "普米族"
    Case 41
        strNation = "塔吉克族"
    Case 42
        strNation = "怒族"
    Case 43
        strNation = "乌孜别克族"
    Case 44
        strNation = "俄罗斯族"
    Case 45
        strNation = "鄂温克族"
    Case 46
        strNation = "德昴族"
    Case 47
        strNation = "保安族"
    Case 48
        strNation = "裕固族"
    Case 49
        strNation = "京族"
    Case 50
        strNation = "塔塔尔族"
    Case 51
        strNation = "独龙族"
    Case 52
        strNation = "鄂伦春族"
    Case 53
        strNation = "赫哲族"
    Case 54
        strNation = "门巴族"
    Case 55
        strNation = "珞巴族"
    Case 56
        strNation = "基诺族"
    Case Else
        strNation = "其他"
    End Select
    TranNation = strNation
End Function



