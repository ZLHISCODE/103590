VERSION 5.00
Begin VB.Form frmStyle_SingleMan 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   -60
   ClientTop       =   -45
   ClientWidth     =   11955
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmrRefreshInterval 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7560
      Top             =   240
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   6600
      Top             =   240
   End
   Begin VB.Label lblClinicName1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   8280
      TabIndex        =   6
      Top             =   1680
      Width           =   240
   End
   Begin VB.Label lblPatientInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0002F6FC&
      Height          =   435
      Left            =   1920
      TabIndex        =   5
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image imgDoctor 
      Height          =   1215
      Left            =   7320
      Picture         =   "frmStyle_SingleMan.frx":0000
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1035
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2014年01月19日"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   9600
      TabIndex        =   0
      Top             =   600
      Width           =   1470
   End
   Begin VB.Label lblWeek 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "星期日"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   10080
      TabIndex        =   1
      Top             =   120
      Width           =   990
   End
   Begin VB.Label lblClinicName0 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   7920
      TabIndex        =   2
      Top             =   1320
      Width           =   240
   End
   Begin VB.Label lblDoctorJob 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   8640
      TabIndex        =   4
      Top             =   3480
      Width           =   240
   End
   Begin VB.Label lblDoctorName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   8760
      TabIndex        =   3
      Top             =   2880
      Width           =   240
   End
   Begin VB.Image imgBack 
      Height          =   7215
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmStyle_SingleMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ISty

'需要实现的接口方法如下：
'
'
'打开lcd显示界面
'public sub ISty_Show(byval lngWindowNo as long)
'lngWindowNo:窗口编号，根据窗口编号读取配置信息，并进行显示
'
'end sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private mlngWindowNo As Long            '窗口编号
Private mlngRefreshInterval As Long     '轮询时间间隔
Private mlngInterval As Long            '累计时间间隔
Private mstrStyleTylePath As String     '窗口样式图片路径
Private mstrClinicNames As String       '临床排队业务下的诊室名称
Private mLcdCommonParameter As TLcdCommonParameter

Private Type TPageObj
    tpBackImage     As TRect        '背景(皮肤)

    tpClinicName    As TRect     '诊室名称
    tpDoctorPhoto   As TRect     '医生照片
    tpDoctorName    As TRect     '医生姓名
    tpDoctorJob     As TRect     '医生职称
    tpPatientInfo   As TRect     '病人信息
    tpWeek          As TRect     '星期
    tpDate          As TRect     '日期
End Type

Private mtpPageObj As TPageObj

Private Sub GetSkinObj(ByVal strSkinName As String)
'读取样式配置文件，对界面控件位置进行初始化
    
    Call SetIniFile(strSkinName)
    
    With mtpPageObj
        '背景图大小
        .tpBackImage.lngWidth = Val(ReadValue("皮肤分辨率", "宽"))
        .tpBackImage.lngHeight = Val(ReadValue("皮肤分辨率", "高"))
        
        '诊室名称
        .tpClinicName.lngLeft = Val(ReadValue("诊室名称", "左"))
        .tpClinicName.lngTop = Val(ReadValue("诊室名称", "顶"))
        .tpClinicName.lngWidth = Val(ReadValue("诊室名称", "宽"))
        .tpClinicName.lngHeight = Val(ReadValue("诊室名称", "高"))
        
        '医生照片
        .tpDoctorPhoto.lngLeft = Val(ReadValue("医生照片", "左"))
        .tpDoctorPhoto.lngTop = Val(ReadValue("医生照片", "顶"))
        .tpDoctorPhoto.lngWidth = Val(ReadValue("医生照片", "宽"))
        .tpDoctorPhoto.lngHeight = Val(ReadValue("医生照片", "高"))
        
        '医生姓名
        .tpDoctorName.lngLeft = Val(ReadValue("医生姓名", "左"))
        .tpDoctorName.lngTop = Val(ReadValue("医生姓名", "顶"))
        .tpDoctorName.lngWidth = Val(ReadValue("医生姓名", "宽"))
        .tpDoctorName.lngHeight = Val(ReadValue("医生姓名", "高"))
        
        '医生职称
        .tpDoctorJob.lngLeft = Val(ReadValue("医生职称", "左"))
        .tpDoctorJob.lngTop = Val(ReadValue("医生职称", "顶"))
        .tpDoctorJob.lngWidth = Val(ReadValue("医生职称", "宽"))
        .tpDoctorJob.lngHeight = Val(ReadValue("医生职称", "高"))
        
        '病人信息
        .tpPatientInfo.lngLeft = Val(ReadValue("病人信息", "左"))
        .tpPatientInfo.lngTop = Val(ReadValue("病人信息", "顶"))
        .tpPatientInfo.lngWidth = Val(ReadValue("病人信息", "宽"))
        .tpPatientInfo.lngHeight = Val(ReadValue("病人信息", "高"))
        
        '星期区域
        .tpWeek.lngLeft = Val(ReadValue("星期区域", "左"))
        .tpWeek.lngTop = Val(ReadValue("星期区域", "顶"))
        .tpWeek.lngWidth = Val(ReadValue("星期区域", "宽"))
        .tpWeek.lngHeight = Val(ReadValue("星期区域", "高"))
        
        '日期区域
        .tpDate.lngLeft = Val(ReadValue("日期区域", "左"))
        .tpDate.lngTop = Val(ReadValue("日期区域", "顶"))
        .tpDate.lngWidth = Val(ReadValue("日期区域", "宽"))
        .tpDate.lngHeight = Val(ReadValue("日期区域", "高"))
    End With
End Sub

Public Sub ISty_RefreshQueueData(Optional ByVal lngQueueId As Long)
'刷新界面显示数据
    Call LoadCallingData
    'Call SetStyleFont
    
    '数据刷新后将计时器清0
    mlngInterval = 0
End Sub

'打开lcd显示界面
Public Sub ISty_Show(ByVal lngWindowNo As Long)
'lngWindowNo:窗口编号，根据窗口编号读取配置信息，并进行显示
    mlngWindowNo = lngWindowNo
    
    Call InitMonitor    '初始化监视器设置
    
    Call InitLocalPars
    
    Call LoadCallingData
    
    Call SetStyleFont

    Call Show
End Sub

Public Function ISty_ShowCfg(ByVal lngWindowNo As Long, objOwner As Object) As Boolean
'打开对应的样式配置窗口
    Dim objConfig As frmStyle_CommonCfg
    
    Set objConfig = New frmStyle_CommonCfg
            
    ISty_ShowCfg = objConfig.OpenShowConfig(lngWindowNo, TShowStyle.ssSingleMan, Me)
End Function


Public Function ISty_MsgProcess(ByVal lngWindowNo As Long, _
    ByVal strMsgKey As String, ByVal strXmlContext As String, rsData As ADODB.Recordset) As Boolean
'消息接收处理

    Dim strValue As String

On Error GoTo ErrorHand
    
    '判断消息中的队列名称是否需要进行处理的队列名称
    rsData.Filter = "node_name='queue_name'"
    If rsData.RecordCount <= 0 Then
        Debug.Print "消息无效，检测到未包含有效的队列名称，终止消息处理。"
        Exit Function
    End If

    strValue = Nvl(rsData!node_value)

    If InStr(mLcdCommonParameter.strQueryQueueNames, strValue) <= 0 Then
        Debug.Print "该消息所属队列不属于当前业务处理范围，忽略消息处理。"
        Exit Function
    End If
    
    '根据接收到的消息进行处理......
    Select Case strMsgKey
        Case G_STR_MSG_QUEUE_001, G_STR_MSG_QUEUE_002, G_STR_MSG_QUEUE_003
            Call ISty_RefreshQueueData
    End Select

    Exit Function
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Function


Public Function ISty_WindNo() As Long
'获取当前样式窗口的编号
    ISty_WindNo = mlngWindowNo
End Function


Private Function InitLocalPars() As Boolean
'初始化本地参数设置
    Dim i As Integer
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    Dim lngCurLCDNo As Long
    Dim strBusinessType As String
    Dim strLCDLocation As String
    Dim strQueryQueueNames As String

On Error GoTo ErrorHand
    If gobjFile.FolderExists(App.Path & "\Skin\单病人样式") Then
        mstrStyleTylePath = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "皮肤样式", App.Path & "\Skin\单病人样式\单病人宽屏深蓝") & ".jpg"
    Else
        mstrStyleTylePath = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "皮肤样式", App.Path & "\zlQueueShow\Skin\单病人样式\单病人宽屏深蓝") & ".jpg"
    End If
    
    imgBack.Picture = LoadPicture(mstrStyleTylePath)
    
    Call GetSkinObj(Replace(mstrStyleTylePath, ".jpg", ".ini"))
    
    '显示器编号
    lngCurLCDNo = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "显示器编号", 1)) - 1
    If lngCurLCDNo < 0 Then lngCurLCDNo = 0
        
    '显示模式,0-全屏；1-自定义
    If Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "显示模式", 0)) = 0 Then
        Call SetFullScreenWindow(Me, lngCurLCDNo)
    Else
        strLCDLocation = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "自定义位置")
        
        If strLCDLocation <> "" Then
            mLcdCommonParameter.recPos.lngLeft = Mid(Split(strLCDLocation, "|")(0), 3) * Screen.TwipsPerPixelX
            mLcdCommonParameter.recPos.lngTop = Mid(Split(strLCDLocation, "|")(1), 3) * Screen.TwipsPerPixelY
            mLcdCommonParameter.recPos.lngWidth = Mid(Split(strLCDLocation, "|")(2), 3) * Screen.TwipsPerPixelX
            mLcdCommonParameter.recPos.lngHeight = Mid(Split(strLCDLocation, "|")(3), 3) * Screen.TwipsPerPixelY
        End If
        
        Call SetCustomWindow(Me, lngCurLCDNo, mLcdCommonParameter.recPos)
    End If

    '排队列表中显示的队列名
    strQueryQueueNames = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "显示队列", "")
    
    mLcdCommonParameter.blnConvertQueueName = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "转换队列名称", 0)) = 1
    
    If strQueryQueueNames <> "" Then
        If mLcdCommonParameter.blnConvertQueueName Then    '转换成老版本格式的队列名称
            Select Case glngBusinessType
                Case TBusinessType.btClinical
                    mLcdCommonParameter.strQueryQueueNames = Split(Split(strQueryQueueNames, "|")(1), "_")(0)
                    mstrClinicNames = Split(strQueryQueueNames, ":")(1)
                    
                Case TBusinessType.btPacs
                    For i = 0 To UBound(Split(strQueryQueueNames, ","))
                        If InStr(strQueryQueueNames, "科室队列") > 0 Then    '按科室排队
                            mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & Split(Split(Split(Split(strQueryQueueNames, ",")(i), "|")(1), ":")(0), "_")(1) & "-" & Split(Split(strQueryQueueNames, ",")(i), "|")(0)
                        Else
                            mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & Split(Split(Split(strQueryQueueNames, ",")(i), "|")(1), "_")(0) & ":" & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                        End If
                    Next
                    
                    mLcdCommonParameter.strQueryQueueNames = Mid(mLcdCommonParameter.strQueryQueueNames, 2)
                    
                Case TBusinessType.btPeis
                    strSql = "select 站点名称 from 体检站点分布 where 执行科室id=[1]"
                    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取站点名称", Val(Split(Split(strQueryQueueNames, "|")(1), "_")(0)))
                    
                    If rsRecord.RecordCount > 0 Then
                        mLcdCommonParameter.strQueryQueueNames = Nvl(rsRecord!站点名称) & ":" & Split(strQueryQueueNames, ":")(1)
                    End If
                'case
                '
                '
            End Select
        Else                                                                '''''''''新版队列名称格式
            Select Case glngBusinessType
                Case TBusinessType.btClinical
                    mLcdCommonParameter.strQueryQueueNames = Split(Split(strQueryQueueNames, "|")(1), "_")(0)
                    mstrClinicNames = Split(strQueryQueueNames, ":")(1)
                    
                Case TBusinessType.btPacs
                    For i = 0 To UBound(Split(strQueryQueueNames, ","))
                        mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & Split(Split(strQueryQueueNames, ",")(i), "|")(0) & "-" & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                    Next
                    
                    mLcdCommonParameter.strQueryQueueNames = Mid(mLcdCommonParameter.strQueryQueueNames, 2)
                    
                Case TBusinessType.btPeis
                    strSql = "select 站点名称 from 体检站点分布 where 执行科室id=[1]"
                    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取站点名称", Val(Split(Split(strQueryQueueNames, "|")(1), "_")(0)))
                    
                    If rsRecord.RecordCount > 0 Then
                        mLcdCommonParameter.strQueryQueueNames = Nvl(rsRecord!站点名称) & ":" & Split(strQueryQueueNames, ":")(1)
                    End If
                'case
                '
                '
            End Select
        End If
    End If
    
    '当前诊室名称
    mLcdCommonParameter.strCurDiagnoseRoom = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "本机执行间", "")
    If mLcdCommonParameter.strCurDiagnoseRoom = "" Then
        strSql = "select d.名称 from 上机人员表 A,人员表 B,部门人员 C,部门表 D " & _
                 "where A.人员ID=B.ID And b.id=c.人员id and c.部门id=d.id and c.缺省=1 and A.用户名=[1]"
        
        Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", UCase(gstrUserName))
        
        If rsRecord.RecordCount > 0 Then lblClinicName0.Caption = Nvl(rsRecord!名称)
    Else
        If InStr(strQueryQueueNames, "科室队列") > 0 Then
            lblClinicName0.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(0)
        Else
            If Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "诊室标题是否显示科室名", 0)) = 1 Then
                If Trim(Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)) <> "" Then
                    lblClinicName0.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(0) & Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)
                Else
                    lblClinicName0.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(0)
                End If
            Else
                If Trim(Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)) <> "" Then
                    lblClinicName0.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)
                Else
                    lblClinicName0.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(0)
                End If
            End If
        End If
    End If
    
    If Len(lblClinicName0.Caption) <= 5 Then lblClinicName0 = FormatStr(lblClinicName0.Caption)
    
    lblClinicName1.Caption = lblClinicName0.Caption
    
    mLcdCommonParameter.blnFontAutoSizeToList = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "列表字体自适应", True)
    
    '排队列表轮询间隔
    mlngRefreshInterval = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "轮询间隔", 30))

    Call LoadDoctorInfo
    
    tmrRefreshInterval.Enabled = True
Exit Function
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Function

Private Function FormatStr(ByVal strSources As String) As String
'功能：将字符串中的汉字之间加上空格
    Dim i As Integer
    Dim strResult As String
    Dim strCurS As String
    Dim strNextS As String
    
    If Len(strSources) <= 1 Then Exit Function
    
    For i = 1 To Len(strSources) - 1
        strCurS = Mid(strSources, i, 1)
        strNextS = Mid(strSources, i + 1, 1)
        strResult = strResult & strCurS
        
        If Not (Asc(strNextS) < 255 And Asc(strNextS) > 0) Then
            strResult = strResult & " "
        End If
    Next
    
    FormatStr = strResult & strNextS
End Function

Private Sub LoadDoctorInfo()
'加载对应执行间的医生和科室相关信息
    Dim i As Integer
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    Dim blnNotWorkingTime As Boolean

    Dim strDoctorInfo As String     '保存格式："医生1的姓名和职位|医生2的姓名和职位|。。。。"
    Dim strDoctorPhoto As String    '保存格式："医生1的照片|医生2的照片|。。。。"
    Dim strIntroduction As String   '保存格式："医生1的简介|医生2的简介|。。。。"
    Dim strWorkingTime As String    '保存格式："医生1的值班时间|医生2的值班时间|。。。。"
    
    blnNotWorkingTime = True
    
    strDoctorInfo = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "医生信息")    '
    strDoctorPhoto = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "医生照片")    '
    strWorkingTime = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "值班时间")   '
    strIntroduction = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "医生简介")   '
    
    '根据当前星期读取对应的医生配置信息
    For i = 0 To UBound(Split(Mid(strWorkingTime, 2), "|"))
        If Split(Mid(strWorkingTime, 2), "|")(i) = lblWeek.Caption Then
            blnNotWorkingTime = False
            
            lblDoctorName.Caption = Split(Split(Mid(strDoctorInfo, 2), "|")(i), "-")(1)
            
            If lblDoctorName.Caption <> "" Then
                strSql = "select 专业技术职务 from 人员表 where id=[1]"
                Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", Val(Split(Split(Mid(strDoctorInfo, 2), "|")(i), "-")(0)))
                
                If rsRecord.RecordCount > 0 Then
                    lblDoctorJob.Caption = Nvl(rsRecord!专业技术职务)
                End If
            End If
            
            Call LoadPictureInfo(imgDoctor, Split(Mid(strDoctorPhoto, 2), "|")(i))
            
            Exit Sub
        End If
    Next
    
    '如果只配置医生，没有指定医生的值班信息，则读取登陆人员信息
    If blnNotWorkingTime Then
        strSql = "select B.姓名,B.ID from 上机人员表 A,人员表 B where A.人员ID=B.ID And A.用户名=[1]"
        Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", UCase(gstrUserName))

        If rsRecord.RecordCount > 0 Then
            For i = 0 To UBound(Split(Mid(strDoctorInfo, 2), "|"))
                If Split(Split(Mid(strDoctorInfo, 2), "|")(i), "-")(0) = Nvl(rsRecord!姓名) Then
                    lblDoctorName.Caption = Split(Split(Mid(strDoctorInfo, 2), "|")(i), "-")(1)
                    
                    If lblDoctorName.Caption <> "" Then
                        strSql = "select 专业技术职务 from 人员表 where id=[1]"
                        Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", Split(Split(Mid(strDoctorInfo, 2), "|")(i), "-")(0))
                        
                        If rsRecord.RecordCount > 0 Then
                            lblDoctorJob.Caption = Nvl(rsRecord!专业技术职务)
                        End If
                    End If
                    
                    Call LoadPictureInfo(imgDoctor, Split(Mid(strDoctorPhoto, 2), "|")(i))

                    Exit Sub
                End If
            Next
        End If
    End If
End Sub

Private Sub SetStyleFont()
'设置界面各控件字体属性
    '设置界面各控件字体属性
    Dim i As Integer
    Dim strFontPropertys As String           '格式:"字体:宋体|字号:20|粗体:FALSE|斜体:FALSE|下划线:FALSE|前景色:16777215"
    Dim strFontProperty() As String
On Error GoTo ErrorHand

    '诊室名称
    strFontPropertys = Trim(ReadValue("字体设置", "诊室名称字体", "字体:微软雅黑|字号:50|粗体:TRUE|斜体:FALSE|下划线:FALSE|前景色:194300"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblClinicName0, strFontProperty)
        Call SetControlFont(lblClinicName1, strFontProperty)
        lblClinicName0.ForeColor = vbBlack
    End If
    
    '医生姓名、职称
    strFontPropertys = Trim(ReadValue("字体设置", "医生信息字体", "字体:微软雅黑|字号:22|粗体:TRUE|斜体:FALSE|下划线:FALSE|前景色:0"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblDoctorName, strFontProperty)
        Call SetControlFont(lblDoctorJob, strFontProperty)
    End If
    
    '病人信息字体
    strFontPropertys = Trim(ReadValue("字体设置", "病人信息字体", "字体:微软雅黑|字号:70|粗体:TRUE|斜体:FALSE|下划线:FALSE|前景色:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblPatientInfo, strFontProperty)
    End If
    
    '星期
    strFontPropertys = Trim(ReadValue("字体设置", "星期字体", "字体:微软雅黑|字号:15|粗体:FALSE|斜体:FALSE|下划线:FALSE|前景色:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblWeek, strFontProperty)
    End If

    '日期
    strFontPropertys = Trim(ReadValue("字体设置", "日期字体", "字体:微软雅黑|字号:15|粗体:FALSE|斜体:FALSE|下划线:FALSE|前景色:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblDate, strFontProperty)
    End If
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub LoadCallingData()
'加载处于呼叫中的数据
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    Dim dblHeightScale As Double, dblWidhtScale As Double
    
On Error GoTo ErrorHand:
    lblPatientInfo.Caption = ""
    
    If mLcdCommonParameter.strQueryQueueNames = "" Or glngBusinessType < 0 Then Exit Sub
    
    Select Case glngBusinessType
        Case TBusinessType.btClinical
            strSql = "select 排队号码,患者姓名,排队状态,呼叫时间,备注,排队序号,诊室 from 排队叫号队列 where 队列名称=[1] and 诊室=[2] and 业务类型=[3] and " & _
                     "排队状态 in (1,7,8,9) And 排队时间 > trunc(sysdate) "
            
            Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取排队信息", mLcdCommonParameter.strQueryQueueNames, mstrClinicNames, glngBusinessType)
            
        Case TBusinessType.btPacs
            strSql = "select 排队号码,患者姓名,排队状态,呼叫时间,备注,排队序号,诊室 from 排队叫号队列 where 队列名称=[1] and 诊室=[2] and 业务类型=[3] and " & _
                     "排队状态 in (1,7,8,9) And 排队时间 > trunc(sysdate) "
            
            If mLcdCommonParameter.strCurDiagnoseRoom = "" Then Exit Sub
            Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取排队信息", mLcdCommonParameter.strQueryQueueNames, CStr(Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)), glngBusinessType)
            
        Case TBusinessType.btPeis
            strSql = "select 排队号码,患者姓名,排队状态,呼叫时间,备注,排队序号,诊室 from 排队叫号队列 where 队列名称=[1] and 业务类型=[2] and " & _
                     "排队状态 in (1,7,8,9) And 排队时间 > trunc(sysdate) "
            
            Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取排队信息", mLcdCommonParameter.strQueryQueueNames, glngBusinessType)
         
        'case
        '
        '
    End Select

    If rsRecord.RecordCount >= 1 Then rsRecord.Filter = "排队状态=9"
    If rsRecord.RecordCount <= 0 Then rsRecord.Filter = "排队状态=1"
    If rsRecord.RecordCount <= 0 Then rsRecord.Filter = "排队状态=7"
    If rsRecord.RecordCount <= 0 Then rsRecord.Filter = "排队状态=8"
    If rsRecord.RecordCount <= 0 Then Exit Sub
    
    rsRecord.Sort = "呼叫时间 desc"
    lblPatientInfo.Caption = Format(Nvl(rsRecord!排队号码), "000") & "号   " & Nvl(rsRecord!患者姓名) & IIf(Len(Trim(Nvl(rsRecord!患者姓名))) <= 3, "   ", "")
    
    '重新设置当前呼叫信息显示位置和字体大小
    dblHeightScale = imgBack.Height / mtpPageObj.tpBackImage.lngHeight
    dblWidhtScale = imgBack.Width / mtpPageObj.tpBackImage.lngWidth
    
    If mLcdCommonParameter.blnFontAutoSizeToList Then
        lblPatientInfo.Caption = Trim(lblPatientInfo.Caption)
        
        lblPatientInfo.Visible = False
        lblPatientInfo.FontSize = dblHeightScale * mtpPageObj.tpPatientInfo.lngHeight / 20

        While lblPatientInfo.Width > dblHeightScale * mtpPageObj.tpPatientInfo.lngWidth Or lblPatientInfo.Height > dblHeightScale * mtpPageObj.tpPatientInfo.lngHeight
            lblPatientInfo.FontSize = lblPatientInfo.FontSize - 1
        Wend
        lblPatientInfo.Visible = True
    End If
    
    lblPatientInfo.Left = dblWidhtScale * mtpPageObj.tpPatientInfo.lngLeft + dblWidhtScale * mtpPageObj.tpPatientInfo.lngWidth / 2 - lblPatientInfo.Width / 2
    lblPatientInfo.Top = dblHeightScale * mtpPageObj.tpPatientInfo.lngTop + dblHeightScale * mtpPageObj.tpPatientInfo.lngHeight / 2 - lblPatientInfo.Height / 2
Exit Sub
ErrorHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHand
    mlngInterval = 0
    tmrRefreshInterval.Interval = 1000
    
    Call refreshWeekLab
    
    lblDate.Caption = Format(Now, "yyyy年mm月dd日 hh:mm:ss")
Exit Sub
ErrorHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    Dim dblHeightScale As Double, dblWidhtScale As Double
    
    '窗体背景
    imgBack.Left = 0
    imgBack.Top = 0
    imgBack.Height = Me.ScaleHeight
    imgBack.Width = Me.ScaleWidth
    
    dblHeightScale = imgBack.Height / mtpPageObj.tpBackImage.lngHeight
    dblWidhtScale = imgBack.Width / mtpPageObj.tpBackImage.lngWidth

    '诊室名称
    lblClinicName0.Left = dblWidhtScale * mtpPageObj.tpClinicName.lngLeft + dblWidhtScale * mtpPageObj.tpClinicName.lngWidth / 2 - lblClinicName0.Width / 2
    lblClinicName0.Top = dblHeightScale * mtpPageObj.tpClinicName.lngTop + dblHeightScale * mtpPageObj.tpClinicName.lngHeight / 2 - lblClinicName0.Height / 2
    
    lblClinicName1.Left = lblClinicName0.Left - 50
    lblClinicName1.Top = lblClinicName0.Top - 50
    '医生照片
    Call ResizeImg(imgDoctor, dblWidhtScale * mtpPageObj.tpDoctorPhoto.lngLeft, dblHeightScale * mtpPageObj.tpDoctorPhoto.lngTop, dblWidhtScale * mtpPageObj.tpDoctorPhoto.lngWidth, dblHeightScale * mtpPageObj.tpDoctorPhoto.lngHeight)

    '医生姓名
    lblDoctorName.Left = dblWidhtScale * mtpPageObj.tpDoctorName.lngLeft + dblWidhtScale * mtpPageObj.tpDoctorName.lngWidth / 2 - lblDoctorName.Width / 2
    lblDoctorName.Top = dblHeightScale * mtpPageObj.tpDoctorName.lngTop + dblHeightScale * mtpPageObj.tpDoctorName.lngHeight / 2 - lblDoctorName.Height / 2
    
    '医生职位
    lblDoctorJob.Left = dblWidhtScale * mtpPageObj.tpDoctorJob.lngLeft + dblWidhtScale * mtpPageObj.tpDoctorJob.lngWidth / 2 - lblDoctorJob.Width / 2
    lblDoctorJob.Top = dblHeightScale * mtpPageObj.tpDoctorJob.lngTop + dblHeightScale * mtpPageObj.tpDoctorJob.lngHeight / 2 - lblDoctorJob.Height / 2
    
    '病人信息
    If mLcdCommonParameter.blnFontAutoSizeToList Then
        lblPatientInfo.Caption = Trim(lblPatientInfo.Caption)
        
        lblPatientInfo.Visible = False
        lblPatientInfo.FontSize = dblHeightScale * mtpPageObj.tpPatientInfo.lngHeight / 20

        While lblPatientInfo.Width > dblHeightScale * mtpPageObj.tpPatientInfo.lngWidth Or lblPatientInfo.Height > dblHeightScale * mtpPageObj.tpPatientInfo.lngHeight
            lblPatientInfo.FontSize = lblPatientInfo.FontSize - 1
        Wend
        lblPatientInfo.Visible = True
    End If
    
    lblPatientInfo.Left = dblWidhtScale * mtpPageObj.tpPatientInfo.lngLeft + dblWidhtScale * mtpPageObj.tpPatientInfo.lngWidth / 2 - lblPatientInfo.Width / 2
    lblPatientInfo.Top = dblHeightScale * mtpPageObj.tpPatientInfo.lngTop + dblHeightScale * mtpPageObj.tpPatientInfo.lngHeight / 2 - lblPatientInfo.Height / 2
    
    '日期
    lblDate.Left = dblWidhtScale * mtpPageObj.tpDate.lngLeft + dblWidhtScale * mtpPageObj.tpDate.lngWidth / 2 - lblDate.Width / 2
    lblDate.Top = dblHeightScale * mtpPageObj.tpDate.lngTop + dblHeightScale * mtpPageObj.tpDate.lngHeight / 2 - lblDate.Height / 2
    
    '星期
    lblWeek.Left = dblWidhtScale * mtpPageObj.tpWeek.lngLeft + dblWidhtScale * mtpPageObj.tpWeek.lngWidth / 2 - lblWeek.Width / 2
    lblWeek.Top = dblHeightScale * mtpPageObj.tpWeek.lngTop + dblHeightScale * mtpPageObj.tpWeek.lngHeight / 2 - lblWeek.Height / 2
End Sub

Private Sub tmrRefreshInterval_Timer()
On Error GoTo ErrorHand
    mlngInterval = mlngInterval + 1

    '当Timer累计的时间小于轮询时间时，无需刷新排队数据
    If mlngInterval < mlngRefreshInterval Then Exit Sub
    '将累计的时间清0
    mlngInterval = 0
    
    Call LoadCallingData
'    Call SetStyleFont
Exit Sub
ErrorHand:
    Debug.Print Err.Description
    Err.Clear
End Sub

Private Sub tmrTime_Timer()
On Error GoTo ErrorHand
    Call refreshWeekLab
    lblDate.Caption = Format(Now, "yyyy年mm月dd日 hh:mm:ss")
Exit Sub
ErrorHand:
    Debug.Print Err.Description
    Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHand
    If KeyAscii = vbKeyEscape Then Call CloseStyleWindow
Exit Sub
ErrorHand:
    Debug.Print Err.Description
    Err.Clear
End Sub

Private Sub refreshWeekLab()
    Select Case Weekday(Date)
        Case 1
            lblWeek.Caption = "星期日"
        Case 2
            lblWeek.Caption = "星期一"
        Case 3
            lblWeek.Caption = "星期二"
        Case 4
            lblWeek.Caption = "星期三"
        Case 5
            lblWeek.Caption = "星期四"
        Case 6
            lblWeek.Caption = "星期五"
        Case 7
            lblWeek.Caption = "星期六"
    End Select
End Sub


