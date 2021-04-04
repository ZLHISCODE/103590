VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLedVoice"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit '要求变量声音
Private blnInit_Dev_MDT_SD01 As Boolean  '必须要在退出时关闭设备
Private blnDev_surpass As Boolean

'在模态窗体中显示非模态窗体
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWNOACTIVATE = 4

Private Const HWND_DESKTOP = 0
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub Reset(commTemp As Object)
    '''''''''''''''''''''''''''''''''''''''''''
    '功能：初始参数
    '参数：
    '返回：
    '''''''''''''''''''''''''''''''''''''''''''
    Dim mySys As OSVERSIONINFO
    Dim lngResult As Long
        
    On Error Resume Next
    
    Set ctlComm = commTemp
    '取操作系统
    mySys.dwOSVersionInfoSize = Len(mySys)   '应该是 148
    lngResult = GetVersionEx(mySys)
    
    Select Case mySys.dwPlatformId
        Case VER_PLATFORM_WIN32s
            gintOS = 0
        Case VER_PLATFORM_WIN32_WINDOWS
            gintOS = 1
        Case VER_PLATFORM_WIN32_NT
            gintOS = 2
    End Select
    
    '取注册表中的参数
    gintDevice = Val(GetSetting("ZLSOFT", "公共全局", "设备名称", 1))
    gintPort = Val(GetSetting("ZLSOFT", "公共全局", "端口", 1))
    gstrSpeed = Val(GetSetting("ZLSOFT", "公共全局", "波特率", 9600))
    gblnDDisplay = Val(GetSetting("ZLSOFT", "公共全局", "双屏显示器", 0)) = 1
    
    gblnHaveBottom = Val(GetSetting("ZLSOFT", "公共全局", "有底行信息", 0)) <> 0
    gstrBottom = GetSetting("ZLSOFT", "公共全局", "底行信息", "")
    '曾明春(2005-10-12)
    gblnNewDev = Val(GetSetting("ZLSOFT", "公共全局", "新型SHY-II", 0)) <> 0
    gbln个帐余额 = Val(GetSetting("ZLSOFT", "公共全局", "个帐余额提示", 0)) <> 0
    
    Select Case gintDevice
        Case Dev_SYC_XII
            SycVoice "f"
            'SycVoice "*"
        Case Dev_SYC_Q9
            Call SYC_Q9(gintPort, "*")
        Case Dev_LK822
            With ctlComm
                If .portopen = True Then
                    .portopen = False
                End If
                    
                If .commport = False Then
                    .commport = True
                End If
                '.commport = gintPort
                .Settings = gstrSpeed & "n,8,1"
                .InputLen = 0
                
                .portopen = True
                .output = Chr(27) + "@"
                .output = Chr(27) + "CLR"
            End With
        Case Dev_SHY_II
            SHYVoice "F"
        Case Dev_NJF_VH
            Set gobjLED = CreateObject("CTSVR.Bjq")
            If Not gobjLED Is Nothing Then
                gobjLED.Comport = gintPort
                gobjLED.DispMode = 0
                gobjLED.Display "~ "
                gobjLED.Reset '该命令单独重复执行时很可能会乱屏
                Set gobjLED = Nothing
            End If
        Case Dev_TDKJ_BJ
            Call TDKJ_BJ_FUN(gintPort, "&Sc$")
        Case Dev_TDKJ_BJ_2008
            Call TDKJ_BJ_2008(gintPort, "&Sc$")
        Case Dev_MDT_SD01
                If Not blnInit_Dev_MDT_SD01 Then '多次初始化时,VB会出错并无法控制
                    InitService
                    InitDevice gintPort
                    blnInit_Dev_MDT_SD01 = True
                End If
                Clear_Screen
        Case Dev_surpass
            If Not blnDev_surpass Then '多次初始化时,VB会出错并无法控制
               SetComNo (gintPort)
               blnDev_surpass = True
            End If
        Case Dev_FS_YL01
            Call opencomm(gintPort)
        Case Dev_TDKJ_BJ_IV
            Call TDKJ_BJ_IV(gintPort, "&Sc$")
    End Select
End Sub

Public Sub Init(ByVal strDisplay As String, ByVal lngModule As Long, cnMain As ADODB.Connection)
    '''''''''''''''''''''''''''''''''''''''''''
    '功能：初始化显示                       如：    '''''''''''''''''''''''''
    '                                               '                       '
    '                                               '欢迎到本院就诊         '
    '                                               '     ***收费员为你服务 '
    '                                               '                       '
    '                                               '''''''''''''''''''''''''
    '
    '参数：
    '   strDisplay:待显示的第二行字符     如：***收费员为你服务
    '   lngModule:调用者的模块编号
    '   cnMain:数据库连接对象
    '
    '返回：
    '''''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer
    Dim str请稍等 As String
    
    Select Case gintDevice
        Case Dev_SYC_XII
            SycVoice "$1"
            SycVoice "# 欢迎到本院就诊 #"
            SycVoice "$2"
            SycVoice "#" & strDisplay & "#"
        Case Dev_SYC_Q9
            Call SYC_Q9(gintPort, "&Sc$")
            Call SYC_Q9(gintPort, "# 欢迎到本院就诊 #")
            Call SYC_Q9(gintPort, "$2")
            Call SYC_Q9(gintPort, "#" & strDisplay & "#")
        Case Dev_LK822
            strDisplay = Trim(strDisplay)
            For i = 1 To 15 - Len(strDisplay)
                strDisplay = " " & strDisplay
            Next
            ShowLED "欢迎到本院就诊", " 祝 你 早 日 康 复 ", strDisplay, gstrBottom
        Case Dev_SHY_II
        
        Case Dev_NJF_VH
            On Error Resume Next
            Set gobjLED = CreateObject("CTSVR.Bjq")
            If Not gobjLED Is Nothing Then
                gobjLED.Comport = gintPort
                gobjLED.DispMode = 0
                gobjLED.Display "~" & strDisplay & ".^" & "欢迎到本院就诊,祝您早日康复!"
                Set gobjLED = Nothing
            End If
        Case Dev_TDKJ_BJ
            If strDisplay Like "*收费员为您服务" Then
                strDisplay = Trim(Replace(strDisplay, "收费员为您服务", ""))
            End If
            Call TDKJ_BJ_FUN(gintPort, "&Sc$")
            Call TDKJ_BJ_FUN(gintPort, "&C11 欢迎到本院就诊$")
            Call TDKJ_BJ_FUN(gintPort, "&C21  祝您早日康复$")
            Call TDKJ_BJ_FUN(gintPort, "&C41收费员:" & strDisplay & "$")
        
        Case Dev_TDKJ_BJ_2008
            If strDisplay Like "*收费员为您服务" Then
                strDisplay = Trim(Replace(strDisplay, "收费员为您服务", ""))
            End If
            Call TDKJ_BJ_2008(gintPort, "&Sc$")
            Call TDKJ_BJ_2008(gintPort, "&C11 欢迎到本院就诊$")
            Call TDKJ_BJ_2008(gintPort, "&C21  祝您早日康复$")
            Call TDKJ_BJ_2008(gintPort, "&C41收费员:" & strDisplay & "$")
        Case Dev_MDT_SD01
            If strDisplay Like "*收费员为您服务" Then
                strDisplay = Trim(Replace(strDisplay, "收费员为您服务", ""))
            End If
            Display_Line "你好,欢迎到本院就诊", 4, 0
            Voices "010208"  '你好,欢迎光临,祝你早日康复
            Display_Line SetLength("收费员:" & strDisplay, 20), 4, 1
        Case Dev_surpass
            If strDisplay Like "*收费员为您服务" Then
                strDisplay = Trim(Replace(strDisplay, "收费员为您服务", ""))
            End If
            Call AllClear  '清屏
            str请稍等 = "您好，请稍等.wav"
            Call LocStringDisplay(2, 2, "你好,请稍等" + Chr(0))
            Call LocStringDisplay(2, 22, "祝你早日康复" + Chr(0))
            Call LocStringDisplay(2, 42, "收费员:" & strDisplay & "" + Chr(0))
        Case Dev_FS_YL01
            '
        Case Dev_TDKJ_BJ_IV
            '--加载图片
            Dim rsTmp As New ADODB.Recordset
            Dim strsql As String
            
            Err = 0
            On Error GoTo ErrHand
            If gBlnPic = False Then
                Set rsTmp = zlDatabase.GetUserInfo()
'                strsql = "select 人员id,照片 from 人员照片 where 人员id=[1]"
'                Set rsTmp = zlDatabase.OpenSQLRecord(strsql, "读取照片", Val(rsTmp!ID))
                '90974:李南春，2015/11/25，Lob图片调用方式
                If rsTmp.RecordCount > 0 Then
                    '读取照片
                    zlCommFun.ShowFlash "请稍等，正在加载人员照片……"
'                    Call zlDatabase.ReadPicture(rsTmp, "照片", "c:\" & rsTmp!人员ID & ".bmp")
                    Call sys.Readlob(100, 15, rsTmp!ID, "c:\" & rsTmp!ID & ".bmp")
                    
                    Call TDKJ_BJ_IV(gintPort, "&B0,0,c:\" & rsTmp!ID & ".bmp$")
                    zlCommFun.StopFlash
                End If
                gBlnPic = True
            End If
            Call TDKJ_BJ_IV(gintPort, "&Sc$")
''            Call TDKJ_BJ_IV(gintPort, "&C21 欢迎到本院就诊$")
''            Call TDKJ_BJ_IV(gintPort, "&C31  祝您早日康复$")
''            Call TDKJ_BJ_IV(gintPort, "&C41  收费员:" & strDisplay & "$")
    End Select
    Exit Sub
ErrHand:
    MsgBox "接口调用失败:" & vbCrLf & vbCrLf & Err.Description, vbInformation, gstrSysName
End Sub

Public Sub DisplayPatient(ByVal strPatient As String, Optional ByVal lngPatient As Long)
'功能:显示病人姓名
'参数:strPatient-病人姓名,如果为空,表示第一次调用时显示本窗口空闲,否则传入姓名(如果是有档案的病人,则传入:姓名 性别 年龄)
'参数:lngPatient-有档案的病人才传入,否则为0
    Dim arrName As Variant
    If gblnDDisplay Or gintDevice = Dev_DDisplay Then
        With frmDisplay
            If Not .Visible Then
                SetParent .hWnd, HWND_DESKTOP
                ShowWindow .hWnd, SW_SHOWNOACTIVATE
                '如果将窗口置于顶层，在该窗口显示的情况下，HIS各功能模块管理界面弹出的MsgBox会显示到分屏上，如病人收费管理重打票据时的询问
                'SetWindowPos .hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
            End If
            
            If strPatient = "" Then '显示窗口空闲
                Call .ShowMain(False)
                Call .ShowFee(False)
                
                .Timer1.Enabled = False
                .lblFree.Visible = True
                .lblPause.Visible = False
            Else
                Call .ShowMain(True)
                
                .lblPatient.Caption = strPatient
                Call .ShowFee(False)
            End If
        End With
    End If
    Select Case gintDevice
        Case Dev_FS_YL01
            If strPatient <> "" And strPatient <> "0" Then
                arrName = Split(strPatient, " ")
                Call Dev_FS_YL01_Voice(arrName(0), 0, 0)
            End If
        Case Dev_TDKJ_BJ
            Call TDKJ_BJ_FUN(gintPort, "&Sc$")
            Call TDKJ_BJ_FUN(gintPort, "&C21 欢迎到本院就诊$")
            Call TDKJ_BJ_FUN(gintPort, "&C31  祝您早日康复$")
            Call TDKJ_BJ_FUN(gintPort, "&C41  " & strPatient & "$")
        Case Dev_TDKJ_BJ_IV
            Call TDKJ_BJ_IV(gintPort, "&Sc$")
            Call TDKJ_BJ_IV(gintPort, "&C21 欢迎到本院就诊$")
            Call TDKJ_BJ_IV(gintPort, "&C31  祝您早日康复$")
            Call TDKJ_BJ_IV(gintPort, "&C41  " & strPatient & "$")
    End Select
    
End Sub

Public Sub Display(ByVal strName As String, ByVal strType As String, ByVal strUnit As String, _
    ByVal dblPrice As Double, ByVal dblNumber As Double, ByVal curMoney As Currency)
'功能：LED显示，但不发音
'参数：  strName:项目  strType: 规格   strUnit:计量单位
        'dblPrice:单价 dblNumber:数量  curMoney:金额
 '例: 药名：青霉素 14﹒00元
     '项目：化验费 7﹒00元
   
    On Error Resume Next '可能存在一行显示不下的情况
    Select Case gintDevice
        Case Dev_SYC_XII
            SycVoice "$1"
            SycVoice "#" & "项目：" & Mid(strName, 1, 16) & "#"
            SycVoice "$2"
            SycVoice "#" & "数量：" & dblNumber & " " & "金额：" & Format(curMoney, "0.00") & "元" & "#"
        Case Dev_SYC_Q9
            Call SYC_Q9(gintPort, "$1")
            Call SYC_Q9(gintPort, "#" & "项目：" & Mid(strName, 1, 16) & "#")
            Call SYC_Q9(gintPort, "$2")
            Call SYC_Q9(gintPort, "#" & "数量：" & dblNumber & " " & "金额：" & Format(curMoney, "0.00") & "元" & "#")
        Case Dev_LK822
            If Len(strName) > 15 Then
                strName = Mid(strName, 1, 15)
            End If
            ShowLED "项目：" & strName, "数量：" & dblNumber & "  单位：" & strUnit, "金额：" & Format(curMoney, "0.00") & "元", gstrBottom
        Case Dev_NJF_VH
            
            Set gobjLED = CreateObject("CTSVR.Bjq")
            If Not gobjLED Is Nothing Then
                gobjLED.Comport = gintPort
                gobjLED.DispMode = 0
                If dblNumber = 1 And strUnit = "" Then
                    '简单收费模式
                    gobjLED.Display "~项目:" & strName
                    gobjLED.Display "^金额:" & FormatEx(curMoney, 2) & "元"
                Else
                    gobjLED.Display "~" & strName & "," & FormatEx(dblNumber, 4) & strUnit
                    gobjLED.Display "^" & "单价:" & FormatEx(dblPrice, 3) & "元,金额:" & FormatEx(curMoney, 2) & "元"
                End If
                Set gobjLED = Nothing
            End If
        Case Dev_TDKJ_BJ
            '可以显示4行，每行8个汉字
            Call TDKJ_BJ_FUN(gintPort, "&Sc$") '清屏
            Call TDKJ_BJ_FUN(gintPort, "&C11" & SetLength(strName, 16) & "$")
            Call TDKJ_BJ_FUN(gintPort, "&C21" & SetLength("数量:" & FormatEx(dblNumber, 5) & " " & strUnit, 16) & "$")
            Call TDKJ_BJ_FUN(gintPort, "&C31" & SetLength("单价:" & FormatEx(dblPrice, 5) & " 元", 16) & "$")
            Call TDKJ_BJ_FUN(gintPort, "&C41" & SetLength("金额:" & FormatEx(curMoney, 5) & " 元", 16) & "$")
        Case Dev_TDKJ_BJ_2008
            '可以显示4行，每行8个汉字
            Call TDKJ_BJ_2008(gintPort, "&Sc$") '清屏
            Call TDKJ_BJ_2008(gintPort, "&C11" & SetLength(strName, 16) & "$")
            Call TDKJ_BJ_2008(gintPort, "&C21" & SetLength("数量:" & FormatEx(dblNumber, 5) & " " & strUnit, 16) & "$")
            Call TDKJ_BJ_2008(gintPort, "&C31" & SetLength("单价:" & FormatEx(dblPrice, 5) & " 元", 16) & "$")
            Call TDKJ_BJ_2008(gintPort, "&C41" & SetLength("金额:" & FormatEx(curMoney, 5) & " 元", 16) & "$")
        Case Dev_MDT_SD01
            '可以显示4行,每行10个汉字
            Clear_Screen
            Display_Line SetLength(strName, 20), 0, 0
            Display_Line SetLength("规格:" & strType & " 单位:" & strUnit, 20), 0, 1
            Display_Line SetLength("单价:" & FormatEx(dblPrice, 5) & "元,数量:" & FormatEx(dblNumber, 5), 20), 0, 2
            Display_Line SetLength("金额:" & FormatEx(curMoney, 5) & " 元", 20), 0, 3
        Case Dev_surpass
             '先清屏
            'Call AllClear
            'Call StringDisplay("合计：" & Format(curMoney, "0.00") & "元" + Chr(0), 0, 0)
           ' Call StringDisplay("实收：" & Format(curMoney, "0.00") & "元" + Chr(0), 0, 0)
            'Call StringDisplay("找您：" & Format(curMoney, "0.00") & "元" + Chr(0), 0, 0)
        Case Dev_TDKJ_BJ_IV
''            '可以显示4行，每行8个汉字
            Call TDKJ_BJ_IV(gintPort, "&Sc$") '清屏
''            Call TDKJ_BJ_IV(gintPort, "&C11" & SetLength(strName, 16) & "$")
''            Call TDKJ_BJ_IV(gintPort, "&C21" & SetLength("数量:" & FormatEx(dblNumber, 5) & " " & strUnit, 16) & "$")
''            Call TDKJ_BJ_IV(gintPort, "&C31" & SetLength("单价:" & FormatEx(dblPrice, 5) & " 元", 16) & "$")
''            Call TDKJ_BJ_IV(gintPort, "&C41" & SetLength("金额:" & FormatEx(curMoney, 5) & " 元", 16) & "$")
    End Select
End Sub

Public Sub Speak(ByVal strSpeak As String)
'功能：声音播放，同时LED显示
'参数：
'   strSpeak:
'           #+序号 或 #+序号+空格+金额    表示播放特定信息
'           ~+数字                        表示播放数字
'序号对应功能：
'#0  --请输入密码(单独新增)
'#1  --您好,请稍等
'#2  --谢谢
'#3  --请当面点清, 谢谢!
'#4  --请问您的姓名
'#5  --请您出示磁卡
'#6  --请您到中药房批价
'#7  --请您到X光室批价
'#8  --请您到注射室做皮试
'#9  --请您到门诊办公室审核盖章
'#10 --请您到挂号室输入门诊号
'#11 --请您出示身份证和医保凭证
'#12 --请您出示身份证和公费医疗凭证
'#13 --请您出示医保凭证和公费医疗凭证
'#14 --请问您挂什么科
'#15 --请问您是初诊还是复诊
'#16 --请问您挂专家门诊还是普通门诊
'#17 --请您先预检, 然后再挂号
'#18 --请您先填好病历卡
'#19 --请您出示病历卡
'#20 --请您到B超室批价
'#21 1234.56   --请您付款一千二百三十四点五六元  J
'#22 1234.56   --预收一千二百三十四点五六元 Y
'#23 1234.56   --找零一千二百三十四点五六元 Z

'#24 --请你出示社保卡
'#25 1234.56   --你的费用为XXXX元
'#26 1234.56   --卡上余额XXXX元
'#27 1234.56   --你的卡上余额不足请付现金XXXX元
'#28 --请你先做医保身份鉴别
'#30  请出示就诊卡::32663

'#50 --请您出示医保凭证
'#51 --门诊收费:请输入姓名

    If Len(strSpeak) < 2 Then Exit Sub
    If Left(strSpeak, 1) <> "#" And Left(strSpeak, 1) <> "~" Then Exit Sub
    
    Select Case gintDevice
        Case Dev_SYC_XII
            ContrastSYC_XII strSpeak
        Case Dev_LK822
        
        Case Dev_SHY_II
            ContrastSHY_II strSpeak
        Case Dev_NJF_VH
            Contrast_NJF_VH strSpeak
        Case Dev_TDKJ_BJ
            Dev_TDKJ_BJ_Speak strSpeak
        Case Dev_TDKJ_BJ_2008
            Dev_TDKJ_BJ_2008_Speak strSpeak
        Case Dev_MDT_SD01
            Dev_MDT_SD01_Speak strSpeak
        Case Dev_surpass
            Dev_surpass_speak strSpeak
        Case Dev_FS_YL01
            If InStr(1, strSpeak, "#22 ") > 0 Then
                If IsNumeric(Mid(strSpeak, 5)) Then
                    Call Dev_FS_YL01_Voice(Mid(strSpeak, 5), 1, 0)
                End If
            End If
        Case Dev_TDKJ_BJ_IV
            Dev_TDKJ_BJ_IV_Speak strSpeak
        Case Dev_SYC_Q9
            ContrastSYC_Q9 strSpeak
    End Select
    
    If gblnDDisplay Or gintDevice = Dev_DDisplay Then
        If InStr(1, strSpeak, "#21 ") > 0 Then
        '应缴获得焦点时调用
            With frmDisplay
                .ShowFee (True)
                .lblCash(0).Caption = Replace(strSpeak, "#21 ", "") & "元"
                .lblCash(1).Caption = ""
                .lblCash(2).Caption = ""
                .lblDrugWindow.Caption = ""
            End With
        End If
    End If
End Sub

Public Sub DispCharge(ByVal dbl应收 As Double, ByVal dbl实收 As Double, ByVal dbl找补 As Double)
   '功能:显示应收,实收,找补,并发声
   Dim dbl合计 As Double, dbl找您 As Double
   Dim str应收 As String, str实收 As String, str找您 As String, str找零 As String, str请付款 As String
    Select Case gintDevice
        Case Dev_SYC_XII
            
        Case Dev_SYC_Q9
            
        Case Dev_LK822
            ShowLED "应收：" & Format(dbl应收, "0.00") & "元", "实收：" & Format(dbl实收, "0.00") & "元", "找补：" & Format(dbl找补, "0.00") & "元", gstrBottom
        Case Dev_SHY_II
            
        Case Dev_NJF_VH
            On Error Resume Next
            Set gobjLED = CreateObject("CTSVR.Bjq")
            If Not gobjLED Is Nothing Then
                gobjLED.Comport = gintPort
                gobjLED.DispMode = 0
                gobjLED.Display "~合计:" & Format(dbl应收, "0.00") & "元^收您:" & Format(dbl实收, "0.00") & "元" & IIf(dbl找补 = 0, "", ",找您:" & Format(dbl找补, "0.00") & "元")
                Set gobjLED = Nothing
            End If
        Case Dev_TDKJ_BJ
            '可以显示4行，每行8个汉字
             Call TDKJ_BJ_FUN(gintPort, "&Sc$") '清屏
        Case Dev_TDKJ_BJ_2008
            '可以显示4行，每行8个汉字
             Call TDKJ_BJ_2008(gintPort, "&Sc$")  '清屏
'            Call TDKJ_BJ_FUN(gintPort, "&C11" & SetLength("应收:" & FormatEx(dbl应收, 5), 16) & "$")
'            Call TDKJ_BJ_FUN(gintPort, "&C21" & SetLength("实收:" & FormatEx(dbl实收, 5), 16) & "$")
'            Call TDKJ_BJ_FUN(gintPort, "&C31" & SetLength("找补:" & FormatEx(dbl找补, 5), 16) & "$")
        Case Dev_MDT_SD01
'            Clear_Screen
'            Price CStr(dbl应收)
'            GetPrice CStr(dbl实收)
'            Check CStr(dbl找补)
        Case Dev_surpass
             '先清屏
             Call AllClear
             '显示三行
             dbl合计 = dbl应收
             dbl实收 = dbl实收
             dbl找您 = dbl找补
             str应收 = "应收.wav"
             str实收 = "预收.wav"
             str找零 = "找零.wav"
             str请付款 = "请您付款.wav"
             str找您 = "找零请当面点清谢谢.wav"
             'Call StringDisplay("合计：" & Format(dbl应收, "0.00") & "元" + Chr(0), 0, 0)
             Call LocStringDisplay(2, 2, "应收：" & Format(dbl应收, "0.00") & "元" + Chr(0))
             'Call StringDisplay("实收：" & Format(dbl实收, "0.00") & "元" + Chr(0), 0, 0)
             Call LocStringDisplay(2, 22, "预收：" & Format(dbl实收, "0.00") & "元" + Chr(0))
             'Call StringDisplay("找您：" & Format(dbl找补, "0.00") & "元" + Chr(0), 0, 0)
             Call LocStringDisplay(2, 42, "找零：" & Format(dbl找补, "0.00") & "元" + Chr(0))
             'Call PlayWaves(str找您)
             'Call PlayWaves(App.Path & "\" & inputwave.Text)
'             Call PlayWaves(App.Path & "\" & str请付款)
'             Call RMB2Wav(dbl应收)
'             'Call PlayWaves(App.Path & "\" & str应收)
'             Call PlayWaves(App.Path & "\" & str实收)
'             Call RMB2Wav(dbl实收)
'             Call PlayWaves(App.Path & "\" & str找零)
'             Call RMB2Wav(dbl找补)
'             Call PlayWaves(App.Path & "\" & str找您)
        Case Dev_FS_YL01
            Call Dev_FS_YL01_Voice(dbl应收, 1, IIf(dbl应收 > 1000, 5, 3))
            Call Dev_FS_YL01_Voice(dbl实收, 2, 3)
            Call Dev_FS_YL01_Voice(dbl找补, 3, 0)
        Case Dev_TDKJ_BJ_IV
            '可以显示4行，每行8个汉字
''             Call TDKJ_BJ_IV(gintPort, "&Sc$") '清屏
            Dev_TDKJ_BJ_IV_Speak "#21" & FormatEx(dbl应收, 5)
'            Call TDKJ_BJ_FUN(gintPort, "&C21" & SetLength("实收:" & FormatEx(dbl实收, 5), 16) & "$")
'            Call TDKJ_BJ_FUN(gintPort, "&C31" & SetLength("找补:" & FormatEx(dbl找补, 5), 16) & "$")
    End Select
    
    If gblnDDisplay Or gintDevice = Dev_DDisplay Then
        '显示收费信息
        With frmDisplay
            .ShowFee (True)
            .lblCash(0).Caption = Format(dbl应收, "0.00") & "元"
            .lblCash(1).Caption = Format(dbl实收, "0.00") & "元"
            .lblCash(2).Caption = Format(dbl找补, "0.00") & "元"
            .lblDrugWindow.Caption = ""
            
            .lblThanks.Caption = .lblThanks.Tag
            
            '启用等待计时,可能没有药房信息
            .Timer1.Enabled = True
        End With
    End If
End Sub

Public Sub DisplayBank(ParamArray Bank())
    '功能：在任意行上显示任意字符
    '参数：bank() 为任意个数的参数,如果一个参数都没有，则UBOUND（bank)为一
    '没有返回值
    Dim intBound As Integer, i As Integer, arrTmp As Variant, strTmp As String
        
    intBound = UBound(Bank)
    If intBound = -1 Then Exit Sub
        
    On Error Resume Next
    
    Select Case gintDevice
        Case Dev_SYC_XII
            SycVoice "*"            '清屏
            For i = 0 To intBound
                If i < 2 Then
                    SycVoice "$" & i + 1
                    SycVoice "#" & Bank(i) & "#"
                End If
            Next
        Case Dev_SYC_Q9
            Call SYC_Q9(gintPort, "*")            '清屏)
            For i = 0 To intBound
                If i < 2 Then
                    Call SYC_Q9(gintPort, "$" & i + 1)
                    Call SYC_Q9(gintPort, "#" & Bank(i) & "#")
                End If
            Next
        Case Dev_LK822
            For i = 0 To intBound
                If Len(Bank(i)) > 15 Then
                    Bank(i) = Mid(Bank(i), 1, 15)
                End If
            Next
            If intBound = 0 Then
                ShowLED Bank(0), "", "", gstrBottom
            ElseIf intBound = 1 Then
                ShowLED Bank(0), Bank(1), "", gstrBottom
            ElseIf intBound >= 2 Then
                ShowLED Bank(0), Bank(1), Bank(2), gstrBottom
            End If
        Case Dev_SHY_II
            
        Case Dev_NJF_VH
            Set gobjLED = CreateObject("CTSVR.Bjq")
            If Not gobjLED Is Nothing Then
                gobjLED.Comport = gintPort
                gobjLED.DispMode = 0
                For i = 0 To IIf(intBound > 1, 1, intBound)
                    If i = 0 Then
                        gobjLED.Display "~" & Bank(i)
                    Else
                        gobjLED.Display "^" & Bank(i)
                    End If
                Next
                Set gobjLED = Nothing
            End If
        Case Dev_TDKJ_BJ
            '可以显示4行，每行8个汉字
            Call TDKJ_BJ_FUN(gintPort, "&Sc$") '清屏
            For i = 0 To intBound
                If i < 4 Then
                    Call TDKJ_BJ_FUN(gintPort, "&C" & i + 1 & "1" & SetLength(CStr(Bank(i)), 16) & "$")
                End If
            Next
        Case Dev_TDKJ_BJ_2008
            '可以显示4行，每行8个汉字
            Call TDKJ_BJ_2008(gintPort, "&Sc$")  '清屏
            For i = 0 To intBound
                If i < 4 Then
                    Call TDKJ_BJ_2008(gintPort, "&C" & i + 1 & "1" & SetLength(CStr(Bank(i)), 16) & "$")
                End If
            Next
        Case Dev_MDT_SD01
            Clear_Screen
           For i = 0 To intBound
                If i < 4 Then
                    Display_Line SetLength(CStr(Bank(i)), 20), 0, i
                End If
           Next
'        Case Dev_surpass
             '先清屏
'             Call AllClear
'             For i = 0 To intBound
'                If i < 4 Then
'                    Call StringDisplay("& SetLength(CStr(Bank(i)) &" + Chr(0), 0, 0)
'                End If
'             Next
        Case Dev_TDKJ_BJ_IV
            
    End Select
    
    If gblnDDisplay Or gintDevice = Dev_DDisplay Then
        With frmDisplay
            '显示药房窗口信息
            '收费调用代码:zl9LedVoice.DisplayBank "费用合计:" & txt合计.Text, "取药窗口:" & IIf(mstr西窗 <> "", " " & mstr西窗, "") & IIf(mstr成窗 <> "", " " & mstr成窗, "") & IIf(mstr中窗 <> "", " " & mstr中窗, "")
            If UBound(Bank) >= 1 Then
                If InStr(1, Bank(1), "取药窗口") > 0 Then
                    If .Timer1.Enabled Then .Timer1.Enabled = False
                    If .lblFree.Visible Then .lblFree.Visible = False
                    If .lblPause.Visible Then .lblPause.Visible = False
                    If .lblWaiting.Visible Then .lblWaiting.Visible = False
                    
                    arrTmp = Split(Replace(Bank(1), "取药窗口: ", ""), " ")
                    For i = 0 To UBound(arrTmp)
                        strTmp = strTmp & IIf(strTmp = "", "", vbCrLf) & Choose(i + 1, "门诊西药房", "门诊成药房", "门诊中药房") & arrTmp(i)
                    Next
                    If strTmp = "" Then
                        .lblDrugWindow.Caption = ""
                    Else
                        .lblDrugWindow.Caption = Replace(.lblDrugWindow.Tag, "&Windows", strTmp)
                    End If
                    
                    '启用等待计时
                    .Timer1.Enabled = True
                    Exit Sub
                End If
            End If
            
            '部分支付现金
            '收费调用代码:Call zl9LedVoice.DisplayBank("合计:" & txt合计.Text & "元,应付:" & txt应缴.Text & "元","收您:" & txt缴款.Text & "元" & IIf(Val(txt找补.Text) = 0, "", ",找您:" & txt找补.Text & "元"))
            If UBound(Bank) >= 0 Then
                strTmp = Bank(0)
                If InStr(1, strTmp, "应付") > 0 Then
                    .ShowFee (True)
                    arrTmp = Split(strTmp, ",")
                    .lblCash(0).Caption = Replace(arrTmp(1), "应付:", "")
                    .lblCash(1).Caption = Replace(arrTmp(2), "收您:", "")
                    If UBound(arrTmp) >= 3 Then
                        .lblCash(2).Caption = Replace(arrTmp(3), "找您:", "")
                    Else
                        .lblCash(2).Caption = "0元"
                    End If
                    .lblDrugWindow.Caption = ""
                End If
            End If
        End With
    End If
End Sub

Public Sub VoiceSetting()
    '''''''''''''''''''''''''''''''''''''''''''
    '功能：配置参数设置
    '参数：无参数
    '返回：
    '''''''''''''''''''''''''''''''''''''''''''
    frmSetting.Show vbModal
End Sub

Private Sub Class_Initialize()
    blnInit_Dev_MDT_SD01 = False
    blnDev_surpass = False
End Sub

Private Sub Class_Terminate()
    If blnInit_Dev_MDT_SD01 = True Then
        CloseDevice
        CloseService
    End If
    If blnDev_surpass = True Then
        Call AllClear
        SetComNo (0) '关闭串口
    End If
    If gblnDDisplay Or gintDevice = Dev_DDisplay Then
        '如果本参数设置了显示欢迎信息,会调用本函数,然后再调用displaypatient
         Call Unload(frmDisplay)
    End If
End Sub


