VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmStyle_SingleQueue 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11715
   Icon            =   "frmStyle_SingleQueue.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrRemarkInfo 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8040
      Top             =   120
   End
   Begin VB.Timer tmrTime 
      Interval        =   60000
      Left            =   6240
      Top             =   120
   End
   Begin VB.Timer tmrRefreshInterval 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7200
      Top             =   120
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfCallingData 
      Height          =   2175
      Left            =   1440
      TabIndex        =   7
      Top             =   2280
      Width           =   2895
      _cx             =   5106
      _cy             =   3836
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2141904383
      ForeColor       =   0
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2141904383
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   0
      Rows            =   15
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   0
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Image imgDoctor 
      Height          =   1215
      Left            =   6360
      Picture         =   "frmStyle_SingleQueue.frx":000C
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Label lblDeptInfo 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   6480
      TabIndex        =   11
      Top             =   4920
      Width           =   4215
   End
   Begin VB.Label lblDoctorIntro 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   6480
      TabIndex        =   10
      Top             =   4200
      Width           =   3780
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
      Left            =   8400
      TabIndex        =   9
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
      Left            =   9000
      TabIndex        =   8
      Top             =   2280
      Width           =   240
   End
   Begin VB.Label lblClinicName 
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
      ForeColor       =   &H0000DA3F&
      Height          =   435
      Left            =   840
      TabIndex        =   1
      Top             =   2640
      Width           =   240
   End
   Begin VB.Image imgLOGO 
      Height          =   720
      Left            =   240
      Picture         =   "frmStyle_SingleQueue.frx":13BD
      Stretch         =   -1  'True
      Top             =   60
      Width           =   840
   End
   Begin VB.Label lblWeek 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "星期一"
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
      Left            =   9600
      TabIndex        =   2
      Top             =   0
      Width           =   990
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2013年12月17 14:19"
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
      Left            =   9120
      TabIndex        =   6
      Top             =   480
      Width           =   1890
   End
   Begin VB.Label lblHospitalName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "重庆市第一人民医院"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   4860
   End
   Begin VB.Label lblRemarkInfo 
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
      Left            =   360
      TabIndex        =   4
      Top             =   6480
      Width           =   240
   End
   Begin VB.Label lblCallContext 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "请未叫到号的患者耐心等待！"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6720
      TabIndex        =   3
      Top             =   6600
      Width           =   2340
   End
   Begin VB.Image imgBack 
      Height          =   7335
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmStyle_SingleQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ISty
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                               样式1说明
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'该样式窗口主要是按照指定的一个排队队列进行进行显示
'即到该诊室就诊的对应队列数据
'如该诊室名称为a，假设分配了两个病人分别到a诊室和b诊室，则在这种显示样式下，只显示分配到a诊室的检查
'
'
'目前排队叫号的排队状态分别取值为：
'-1占位（即不正式开始排队）,0-排队中,1-呼叫中,2-已弃号,3-已暂停,4-已完成,7-已呼叫,8-接诊中,9-待呼叫
'排队状态的转换关系如下
'当检查数据进入队列后，默认的排队状态为-1
'当对进入的队列数据执行startqueue方法后，开始正式排队,排队状态被修改为0
'当对该队列数据进行呼叫时，排队状态进入待呼叫状态，即为9
'当语音播放段对该队列播放语音时，排队状态被修改为1
'当语音播放结束后，该队列数据的排队状态被修改为7
'当被呼叫的病人进入该诊室后，在医生选择了接诊操作情况下，该排队状态被修改为8
'当完成该病人的接诊操作后，且医生选择了完成操作，则该排队状态被修改为4
'只有在病人因其他事情需要暂停就诊或者不再进行就诊时，排队状态才被修改为3暂停或者2弃号
'
'在该样式下，需要显示的数据为指定队列下，排队状态为呼叫中，已呼叫和候诊中的数据，候诊即当前正在排队的数据，显示大概形式如下
' 004  张三  呼叫中
'
' 003  王五  已呼叫
' 002  李四  已呼叫
'
' 005  孙六  请候诊
' 006  赵七  请候诊
'
'已呼叫排队检查的显示数量和候诊排队检查的显示数量应可以通过参数设置进行确定。
'已呼叫数据只提取最后几个被呼叫的排队检查数据

'显示样式的效果图片可参考图像文件“样式1”



'需要实现的接口方法如下：
'
'
'打开lcd显示界面
'public sub ISty_Show(byval lngWindowNo as long)
'lngWindowNo:窗口编号，根据窗口编号读取配置信息，并进行显示
'
'end sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const M_LNG_准备就诊显示行 = 2  '排队列表中，准备就诊数据显示行数

Private mlngWindowNo As Long            '窗口编号
Private mlngRefreshInterval As Long     '轮询时间间隔
Private mlngInterval As Long            '累计时间间隔
Private mstrStyleTylePath As String     '窗口样式图片路径
Private mblnShowCallTarget As Boolean   '是否显示就诊目的地
Private mstrClinicNames As String       '临床排队业务下的诊室名称

Private mLcdCommonParameter As TLcdCommonParameter

Private Type TPageObj
    tpBackImage  As TRect        '背景(皮肤)
    tpTopAera    As TRect        '顶部
    tpMiddleAera As TRect        '中部
    tpBottomAera As TRect        '底部
    
    tpHospitalLOGO   As TRect    '医院图标
    tpHospitalName   As TRect    '医院名称
    tpWeek           As TRect    '星期
    tpDate           As TRect    '日期
    
    tpClinicName        As TRect       '诊室名称
    tpDoctorPhotoAera   As TRect       '医生照片区域
    tpDoctorInfo        As TRect       '医生信息
    tpDoctorIntro       As TRect       '医生简介
    tpCurQueuedList     As TRect       '准备候诊列表
    
    tpblnShowListHeader As Boolean      '显示列表标题
    tpstrListHeaderName As String       '列表标题名
    tplngQueueListMaxRows As Long       '单队列列表可以显示的总行数（如果显示列表标题则包括列头）
End Type

Private mtpPageObj As TPageObj

Private Sub GetSkinObj(ByVal strSkinName As String)
'读取样式配置文件，对界面控件位置进行初始化
    
    Call SetIniFile(strSkinName)
    
    With mtpPageObj
        '背景图大小
        .tpBackImage.lngWidth = Val(ReadValue("皮肤分辨率", "宽"))
        .tpBackImage.lngHeight = Val(ReadValue("皮肤分辨率", "高"))
        
        '顶部区域
        .tpTopAera.lngLeft = Val(ReadValue("顶部区域", "左"))
        .tpTopAera.lngTop = Val(ReadValue("顶部区域", "顶"))
        .tpTopAera.lngWidth = Val(ReadValue("顶部区域", "宽"))
        .tpTopAera.lngHeight = Val(ReadValue("顶部区域", "高"))
        
        '医院图标
        .tpHospitalLOGO.lngLeft = Val(ReadValue("医院图标", "左"))
        .tpHospitalLOGO.lngTop = Val(ReadValue("医院图标", "顶"))
        .tpHospitalLOGO.lngWidth = Val(ReadValue("医院图标", "宽"))
        .tpHospitalLOGO.lngHeight = Val(ReadValue("医院图标", "高"))
        
        '医院名称
        .tpHospitalName.lngLeft = Val(ReadValue("医院名称", "左"))
        .tpHospitalName.lngTop = Val(ReadValue("医院名称", "顶"))
        .tpHospitalName.lngWidth = Val(ReadValue("医院名称", "宽"))
        .tpHospitalName.lngHeight = Val(ReadValue("医院名称", "高"))
        
        '星期
        .tpWeek.lngLeft = Val(ReadValue("星期", "左"))
        .tpWeek.lngTop = Val(ReadValue("星期", "顶"))
        .tpWeek.lngWidth = Val(ReadValue("星期", "宽"))
        .tpWeek.lngHeight = Val(ReadValue("星期", "高"))
        
        '日期
        .tpDate.lngLeft = Val(ReadValue("日期", "左"))
        .tpDate.lngTop = Val(ReadValue("日期", "顶"))
        .tpDate.lngWidth = Val(ReadValue("日期", "宽"))
        .tpDate.lngHeight = Val(ReadValue("日期", "高"))
            
        '中部区域
        .tpMiddleAera.lngLeft = Val(ReadValue("中部区域", "左"))
        .tpMiddleAera.lngTop = Val(ReadValue("中部区域", "顶"))
        .tpMiddleAera.lngWidth = Val(ReadValue("中部区域", "宽"))
        .tpMiddleAera.lngHeight = Val(ReadValue("中部区域", "高"))
        
        '底部区域
        .tpBottomAera.lngLeft = Val(ReadValue("底部区域", "左"))
        .tpBottomAera.lngTop = Val(ReadValue("底部区域", "顶"))
        .tpBottomAera.lngWidth = Val(ReadValue("底部区域", "宽"))
        .tpBottomAera.lngHeight = Val(ReadValue("底部区域", "高"))
        
        '排队列表区域
        .tpCurQueuedList.lngLeft = Val(ReadValue("排队列表区域", "左"))
        .tpCurQueuedList.lngTop = Val(ReadValue("排队列表区域", "顶"))
        .tpCurQueuedList.lngWidth = Val(ReadValue("排队列表区域", "宽"))
        .tpCurQueuedList.lngHeight = Val(ReadValue("排队列表区域", "高"))
        
        .tpblnShowListHeader = CBool(ReadValue("排队列表区域", "是否显示列表标题"))
        
        If .tpblnShowListHeader Then
            .tpstrListHeaderName = Trim(ReadValue("排队列表区域", "列表标题名"))
            .tplngQueueListMaxRows = Val(ReadValue("排队列表区域", "总行数")) - 1
        Else
            .tpstrListHeaderName = ""
            .tplngQueueListMaxRows = Val(ReadValue("排队列表区域", "总行数"))
        End If
        
        '诊室名称区域
        .tpClinicName.lngLeft = Val(ReadValue("诊室名称区域", "左"))
        .tpClinicName.lngTop = Val(ReadValue("诊室名称区域", "顶"))
        .tpClinicName.lngWidth = Val(ReadValue("诊室名称区域", "宽"))
        .tpClinicName.lngHeight = Val(ReadValue("诊室名称区域", "高"))
        
        '照片区域
        .tpDoctorPhotoAera.lngLeft = Val(ReadValue("照片区域", "左"))
        .tpDoctorPhotoAera.lngTop = Val(ReadValue("照片区域", "顶"))
        .tpDoctorPhotoAera.lngWidth = Val(ReadValue("照片区域", "宽"))
        .tpDoctorPhotoAera.lngHeight = Val(ReadValue("照片区域", "高"))
        
        '医生信息
        .tpDoctorInfo.lngLeft = Val(ReadValue("医生信息", "左"))
        .tpDoctorInfo.lngTop = Val(ReadValue("医生信息", "顶"))
        .tpDoctorInfo.lngWidth = Val(ReadValue("医生信息", "宽"))
        .tpDoctorInfo.lngHeight = Val(ReadValue("医生信息", "高"))
        
        '医生介绍和科室简介的显示位置
        .tpDoctorIntro.lngLeft = Val(ReadValue("简介区域", "左"))
        .tpDoctorIntro.lngTop = Val(ReadValue("简介区域", "顶"))
        .tpDoctorIntro.lngWidth = Val(ReadValue("简介区域", "宽"))
        .tpDoctorIntro.lngHeight = Val(ReadValue("简介区域", "高"))
    End With
End Sub

Public Sub ISty_RefreshQueueData(Optional ByVal lngQueueId As Long)
'刷新界面显示数据
    Dim blnExist就诊 As Boolean
    
    Call LoadCallingData(blnExist就诊)
    Call SetStyleFont(blnExist就诊)
    
    '数据刷新后将计时器清0
    mlngInterval = 0
End Sub

'打开lcd显示界面
Public Sub ISty_Show(ByVal lngWindowNo As Long)
'lngWindowNo:窗口编号，根据窗口编号读取配置信息，并进行显示
    Dim blnExist就诊 As Boolean     '排队列表中是否存在状态处于“就诊”的数据
    
    mlngWindowNo = lngWindowNo
    
    Call InitMonitor    '初始化监视器设置
    
    If Not InitLocalPars Then Exit Sub
    
    Call LoadCallingData(blnExist就诊)
    
    Call SetStyleFont(blnExist就诊)

    Call Show
End Sub

Public Function ISty_ShowCfg(ByVal lngWindowNo As Long, objOwner As Object) As Boolean
'打开对应的样式配置窗口
    Dim objConfig As frmStyle_CommonCfg
    
    Set objConfig = New frmStyle_CommonCfg
            
    ISty_ShowCfg = objConfig.OpenShowConfig(lngWindowNo, TShowStyle.ssSingleQueue, Me)
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
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    Dim lngCurLCDNo As Long
    Dim strBusinessType As String
    Dim strLCDLocation As String
    Dim strQueryQueueNames As String

On Error GoTo ErrorHand
    If gobjFile.FolderExists(App.Path & "\Skin\单队列样式") Then
        mstrStyleTylePath = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "皮肤样式", App.Path & "\Skin\单队列样式\单队列宽屏深蓝") & ".jpg"
    Else
        mstrStyleTylePath = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "皮肤样式", App.Path & "\zlQueueShow\Skin\单队列样式\单队列宽屏深蓝") & ".jpg"
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

    '数据过滤条件
    mLcdCommonParameter.strFilter = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "过滤条件", "")
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
                    If InStr(strQueryQueueNames, "科室队列") > 0 Then       '按科室排队
                        mLcdCommonParameter.strQueryQueueNames = Split(Split(Split(strQueryQueueNames, "|")(1), ":")(0), "_")(1) & "-" & Split(strQueryQueueNames, "|")(0)
                    Else
                        mLcdCommonParameter.strQueryQueueNames = Split(Split(strQueryQueueNames, "|")(1), "_")(0) & ":" & Split(strQueryQueueNames, ":")(1)
                    End If
                    
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
                    mLcdCommonParameter.strQueryQueueNames = Split(strQueryQueueNames, "|")(0) & "-" & Split(strQueryQueueNames, ":")(1)
                    
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
        
        If rsRecord.RecordCount > 0 Then lblClinicName.Caption = Nvl(rsRecord!名称)
    Else
        If InStr(strQueryQueueNames, "科室队列") > 0 Then
            lblClinicName.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(0)
        Else
            If Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "诊室标题是否显示科室名", 0)) = 1 Then
                If Trim(Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)) <> "" Then
                    lblClinicName.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(0) & Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)
                Else
                    lblClinicName.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(0)
                End If
            Else
                If Trim(Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)) <> "" Then
                    lblClinicName.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)
                Else
                    lblClinicName.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(0)
                End If
            End If
        End If
    End If
    
    '加载医院LOGO
    Call LoadPictureInfo(imgLOGO, GetSetting("ZLSOFT", G_STR_REGPATH, "医院LOGO"))

    '医院名称
    lblHospitalName.Caption = GetSetting("ZLSOFT", G_STR_REGPATH, "医院名称", "重庆市第一人民医院")
    '底端文本
    lblRemarkInfo.Caption = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "底端文本", "")

    '是否滚动显示
    mLcdCommonParameter.blnScrollDisplay = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "滚动显示", "0") = 1
    '排队列表轮询间隔
    mlngRefreshInterval = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "轮询间隔", 30))
    
    '根据显示行设置表格背景图
    If mtpPageObj.tpblnShowListHeader Then
        mLcdCommonParameter.lngQueueRows = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "排队列表显示行", mtpPageObj.tplngQueueListMaxRows)) + 1
        vsfCallingData.WallPaper = CutPicture(mstrStyleTylePath, picTemp, mtpPageObj.tpCurQueuedList.lngLeft * Screen.TwipsPerPixelX, mtpPageObj.tpCurQueuedList.lngTop * Screen.TwipsPerPixelY, mtpPageObj.tpCurQueuedList.lngWidth * Screen.TwipsPerPixelX, mtpPageObj.tpCurQueuedList.lngHeight * Screen.TwipsPerPixelY * mLcdCommonParameter.lngQueueRows / (mtpPageObj.tplngQueueListMaxRows + 1))
    Else
        mLcdCommonParameter.lngQueueRows = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "排队列表显示行", mtpPageObj.tplngQueueListMaxRows))
        vsfCallingData.WallPaper = CutPicture(mstrStyleTylePath, picTemp, mtpPageObj.tpCurQueuedList.lngLeft * Screen.TwipsPerPixelX, mtpPageObj.tpCurQueuedList.lngTop * Screen.TwipsPerPixelY, mtpPageObj.tpCurQueuedList.lngWidth * Screen.TwipsPerPixelX, mtpPageObj.tpCurQueuedList.lngHeight * Screen.TwipsPerPixelY * mLcdCommonParameter.lngQueueRows / mtpPageObj.tplngQueueListMaxRows)
    End If
    
    mLcdCommonParameter.blnFontAutoSizeToList = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "列表字体自适应", True)
    mblnShowCallTarget = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "显示就诊目的地", 0)) = 1
    
    Call LoadDoctorInfo
    
    tmrRefreshInterval.Enabled = True
    tmrRemarkInfo.Enabled = True
    
    InitLocalPars = True
Exit Function
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
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
            
            lblDoctorIntro.Caption = Split(Mid(strIntroduction, 2), "|")(i)
            lblDeptInfo.Visible = False
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
            
                    lblDoctorIntro.Caption = Split(Mid(strIntroduction, 2), "|")(i)
                    lblDeptInfo.Visible = False
                    Exit Sub
                End If
            Next
        End If
    End If
    
    imgDoctor.Visible = False
    lblDoctorName.Visible = False
    lblDoctorJob.Visible = False
    lblDoctorIntro.Visible = False
    lblDoctorIntro.Caption = ""
End Sub

Private Sub SetStyleFont(ByVal blnExist就诊 As Boolean)
'设置界面各控件字体属性
    Dim i As Integer
    Dim strFontPropertys As String           '格式:"字体:宋体|字号:20|粗体:FALSE|斜体:FALSE|下划线:FALSE|前景色:16777215"
    Dim strFontPropertys1 As String
    Dim strFontPropertys2 As String
    Dim strFontPropertys3 As String
    Dim strFontPropertys4 As String
    Dim strFontProperty() As String
    Dim strFontProperty1() As String
    Dim strFontProperty2() As String
    Dim strFontProperty3() As String
    Dim strFontProperty4() As String
    
On Error GoTo ErrorHand
    
    '医生信息字体
    strFontPropertys = Trim(ReadValue("字体设置", "医生信息字体", "字体:宋体|字号:24|粗体:TRUE|斜体:FALSE|下划线:FALSE|前景色:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblDoctorName, strFontProperty)
        Call SetControlFont(lblDoctorJob, strFontProperty)
    End If
    
    '科室/医生简介字体
    strFontPropertys = Trim(ReadValue("字体设置", "医生\科室简介字体", "字体:宋体|字号:15|粗体:TRUE|斜体:FALSE|下划线:FALSE|前景色:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblDoctorIntro, strFontProperty)
        Call SetControlFont(lblDeptInfo, strFontProperty)
    End If
    
    '医院名称
    strFontPropertys = Trim(ReadValue("字体设置", "医院名称字体", "字体:宋体|字号:26|粗体:TRUE|斜体:FALSE|下划线:FALSE|前景色:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblHospitalName, strFontProperty)
    End If
    
    '星期
    strFontPropertys = Trim(ReadValue("字体设置", "星期字体", "字体:宋体|字号:20|粗体:TRUE|斜体:FALSE|下划线:FALSE|前景色:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblWeek, strFontProperty)
    End If

    '日期
    strFontPropertys = Trim(ReadValue("字体设置", "日期字体", "字体:宋体|字号:15|粗体:FALSE|斜体:FALSE|下划线:FALSE|前景色:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblDate, strFontProperty)
    End If

    '诊室名称
    strFontPropertys = Trim(ReadValue("字体设置", "诊室名称字体", "字体:宋体|字号:26|粗体:TRUE|斜体:FALSE|下划线:FALSE|前景色:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblClinicName, strFontProperty)
    End If

   '备注内容
    strFontPropertys = Trim(ReadValue("字体设置", "备注内容字体", "字体:宋体|字号:26|粗体:TRUE|斜体:FALSE|下划线:FALSE|前景色:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblRemarkInfo, strFontProperty)
        Call SetControlFont(lblCallContext, strFontProperty)
    End If
    
    '设置列表字体
    strFontPropertys1 = Trim(ReadValue("字体设置", "排队列表标题字体", "字体:宋体|字号:22|粗体:TRUE|前景色:4471868"))
    strFontPropertys2 = Trim(ReadValue("字体设置", "就诊状态行字体", "字体:宋体|字号:20|粗体:TRUE|前景色:260872"))
    strFontPropertys3 = Trim(ReadValue("字体设置", "准备就诊状态行字体", "字体:宋体|字号:20|粗体:FALSE|前景色:1681613"))
    strFontPropertys4 = Trim(ReadValue("字体设置", "候诊状态行字体", "字体:宋体|字号:20|粗体:FALSE|前景色:16777215"))
    
    strFontProperty1 = Split(strFontPropertys1, "|")
    strFontProperty2 = Split(strFontPropertys2, "|")
    strFontProperty3 = Split(strFontPropertys3, "|")
    strFontProperty4 = Split(strFontPropertys4, "|")
    
    SetVSFListFont vsfCallingData, 0, strFontProperty1
    
    For i = IIf(mtpPageObj.tpblnShowListHeader, 1, 0) To vsfCallingData.Rows - 1
        If InStr(vsfCallingData.TextMatrix(i, 3), "准备就诊") > 0 Then
            SetVSFListFont vsfCallingData, i, strFontProperty3
        ElseIf InStr(vsfCallingData.TextMatrix(i, 3), "就诊") > 0 Then
            SetVSFListFont vsfCallingData, i, strFontProperty2
        ElseIf InStr(vsfCallingData.TextMatrix(i, 3), "候诊") > 0 Then
            SetVSFListFont vsfCallingData, i, strFontProperty4
        End If
    Next
    
    If mLcdCommonParameter.blnFontAutoSizeToList And vsfCallingData.Rows > 1 And vsfCallingData.Cols > 1 Then
        vsfCallingData.Cell(flexcpFontSize, 0, 0, vsfCallingData.Rows - 1, vsfCallingData.Cols - 1) = vsfCallingData.RowHeight(0) / 42
    End If
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub InitDataList(ByVal blnExist优先 As Boolean)
    Dim i As Integer
    Dim lngRow As Long
    '初始化排队列表
    With vsfCallingData
        .Rows = 0   '清空数据
        .BackColorSel = &H838C00
        .ForeColorSel = .ForeColor
        .Cols = IIf(blnExist优先, 5, 4)
        .Rows = mLcdCommonParameter.lngQueueRows
        
        '设置列宽
        If blnExist优先 Then
            .ColWidth(0) = .Width * 1 / 9
            .ColWidth(1) = .Width * 2 / 9
            .ColWidth(2) = .Width * 2 / 9
            .ColWidth(3) = .Width * 2 / 9
            .ColWidth(4) = .Width * 2 / 9
        Else
            .ColWidth(0) = .Width * 1 / 7
            .ColWidth(1) = .Width * 2 / 7
            .ColWidth(2) = .Width * 2 / 7
            .ColWidth(3) = .Width * 2 / 7
        End If
        
        For i = 0 To .Rows - 1
            .RowHeight(i) = .Height / .Rows    '设置行高
        Next
        
        If mtpPageObj.tpblnShowListHeader Then
            .TextMatrix(0, 0) = Trim(ReadValue("排队列表区域", "列表标题名", "排队状态信息"))
            .TextMatrix(0, 1) = .TextMatrix(0, 0)
            .TextMatrix(0, 2) = .TextMatrix(0, 0)
            .TextMatrix(0, 3) = .TextMatrix(0, 0)
            
            If blnExist优先 Then .TextMatrix(0, 4) = .TextMatrix(0, 0)
        
            '标题行合并
            .MergeRow(0) = True
            .MergeCells = flexMergeRestrictRows
            .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        End If
        
        '各行内容显示格式
        For lngRow = IIf(mtpPageObj.tpblnShowListHeader, 1, 0) To .Rows - 1
            .Cell(flexcpAlignment, lngRow, 0) = flexAlignRightCenter
            .Cell(flexcpAlignment, lngRow, 1) = flexAlignCenterCenter
            .Cell(flexcpAlignment, lngRow, 2) = flexAlignLeftCenter
            .Cell(flexcpAlignment, lngRow, 3) = flexAlignLeftCenter
            If blnExist优先 Then .Cell(flexcpAlignment, lngRow, 4) = flexAlignLeftCenter
        Next
    End With
End Sub

Private Sub LoadCallingData(ByRef blnExist就诊 As Boolean)
'加载排队列表数据
'blnExist就诊：当排队列表中有“就诊”时返回true
    Dim i As Integer
    Dim lngRow As Long
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    Dim rsClone As ADODB.Recordset
    Dim strSortStyle As String      '排序方式

On Error GoTo ErrorHand:
    Call InitDataList(False) '排队列表默认为4列,没有“优先原因”列
    
    If mLcdCommonParameter.strQueryQueueNames = "" Or glngBusinessType < 0 Then Exit Sub

    strSql = "select Zl_排队叫号队列_获取排序方式([1]) as 排序方式 from dual"
    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取排序方式", glngBusinessType)
    
    If rsRecord.RecordCount > 0 Then strSortStyle = Nvl(rsRecord!排序方式)
    
    Select Case glngBusinessType
        Case TBusinessType.btClinical
            If mstrClinicNames = "科室队列" Then
                strSql = "select a.排队号码,a.患者姓名,a.排队状态,a.呼叫时间,a.备注,a.排队序号,a.诊室,b.名称 from 排队叫号队列 a,部门表 b where 队列名称=[1] and 业务类型=[2]  and " & _
                         "排队状态 in (0,1,7,8,9) And 排队时间 > trunc(sysdate) and a.科室id=b.id " & IIf(mLcdCommonParameter.strFilter <> "", " and " & mLcdCommonParameter.strFilter, "") & " " & _
                         IIf(strSortStyle <> "", " order by " & strSortStyle, "")
        
                Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取排队信息", mLcdCommonParameter.strQueryQueueNames, glngBusinessType)
            Else
                strSql = "select a.排队号码,a.患者姓名,a.排队状态,a.呼叫时间,a.备注,a.排队序号,a.诊室,b.名称 from 排队叫号队列 a,部门表 b where 队列名称=[1] and (诊室=[2] or 诊室 is null) and 业务类型=[3]  and " & _
                         "排队状态 in (0,1,7,8,9) And 排队时间 > trunc(sysdate) and a.科室id=b.id " & IIf(mLcdCommonParameter.strFilter <> "", " and " & mLcdCommonParameter.strFilter, "") & " " & _
                         IIf(strSortStyle <> "", " order by " & strSortStyle, "")
        
                Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取排队信息", mLcdCommonParameter.strQueryQueueNames, mstrClinicNames, glngBusinessType)
            End If
            
        Case TBusinessType.btPacs
            strSql = "select 排队号码,患者姓名,排队状态,呼叫时间,备注,排队序号,诊室 from 排队叫号队列 where 队列名称=[1] and 业务类型=[2] and " & _
                     "排队状态 in (0,1,7,8,9) And 排队时间 > trunc(sysdate) " & IIf(mLcdCommonParameter.strFilter <> "", " and " & mLcdCommonParameter.strFilter, "") & " " & _
                     IIf(strSortStyle <> "", " order by " & strSortStyle, "")
    
            Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取排队信息", mLcdCommonParameter.strQueryQueueNames, glngBusinessType)
        
        Case TBusinessType.btPeis
            strSql = "select 排队号码,患者姓名,排队状态,呼叫时间,备注,排队序号,诊室 from 排队叫号队列 where 队列名称=[1] and 业务类型=[2] and " & _
                     "排队状态 in (0,1,7,8,9) And 排队时间 > trunc(sysdate) " & IIf(mLcdCommonParameter.strFilter <> "", " and " & mLcdCommonParameter.strFilter, "") & " " & _
                     IIf(strSortStyle <> "", " order by " & strSortStyle, "")
    
            Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取排队信息", mLcdCommonParameter.strQueryQueueNames, glngBusinessType)
        'case
        '
        '
    End Select

    If rsRecord.RecordCount < 1 Then Exit Sub
    Set rsClone = rsRecord.Clone

    rsRecord.Filter = "备注<>"""
    If rsRecord.RecordCount > 0 Then Call InitDataList(True) '有优先原因时，排队列表有5列
    
    If mLcdCommonParameter.blnScrollDisplay Then        '获取处于“呼叫中”和“已呼叫”的数据
        rsRecord.Filter = "排队状态=7"
        rsClone.Filter = "排队状态=1"
        
        lblCallContext.Caption = ""
        
        If rsRecord.RecordCount > 0 Then
            rsRecord.Sort = "呼叫时间 asc"
            rsRecord.MoveFirst
            
            For i = 0 To IIf(rsClone.RecordCount > 0, rsRecord.RecordCount - 1, rsRecord.RecordCount - 2)
                If glngBusinessType = TBusinessType.btClinical Then
                    lblCallContext.Caption = lblCallContext.Caption & " ●" & Format(Nvl(rsRecord!排队号码), "000") & "号 " & Nvl(rsRecord!患者姓名) & " 到 " & IIf(Nvl(rsRecord!诊室) = "", Nvl(rsRecord!名称), Nvl(rsRecord!诊室)) & " 就诊"
                Else
                    lblCallContext.Caption = lblCallContext.Caption & " ●" & Format(Nvl(rsRecord!排队号码), "000") & "号 " & Nvl(rsRecord!患者姓名) & " 到 " & Nvl(rsRecord!诊室) & " 就诊"
                End If
                rsRecord.MoveNext
            Next
        End If
        
        lblRemarkInfo.Caption = ""
        If lblCallContext.Caption = "" Then lblCallContext.Caption = "请未叫到号的患者耐心等待！"
    End If
        
    With vsfCallingData
        rsRecord.Filter = "排队状态=9"  '提取处于“待呼叫”的数据
        If rsRecord.RecordCount <= 0 Then rsRecord.Filter = "排队状态=1" '提取处于“呼叫中”的数据
        
        '当状态处于“呼叫中”没有数据时，提取状态处于“已呼叫”的进行呼叫,当状态处于“已呼叫”的数量不只1条时，提取最近呼叫的一条
        If rsRecord.RecordCount <= 0 Then rsRecord.Filter = "排队状态=7"
        
        '当状态处于“呼叫中”没有数据时，提取状态处于“接诊中”的数据
        If rsRecord.RecordCount <= 0 Then rsRecord.Filter = "排队状态=8"
        
        blnExist就诊 = False
        lngRow = IIf(mtpPageObj.tpblnShowListHeader, 1, 0)
        
        If rsRecord.RecordCount > 0 Then
            rsRecord.Sort = "呼叫时间 desc"

            .TextMatrix(lngRow, 0) = " 请"
            .TextMatrix(lngRow, 1) = Format(rsRecord!排队号码, "000") & "号"
            .TextMatrix(lngRow, 2) = Nvl(rsRecord!患者姓名)
            
            If mblnShowCallTarget Then
                If glngBusinessType = TBusinessType.btClinical Then
                    .TextMatrix(lngRow, 3) = "就诊(" & IIf(Nvl(rsRecord!诊室) = "", Nvl(rsRecord!名称), Nvl(rsRecord!诊室)) & ")"
                Else
                    .TextMatrix(lngRow, 3) = "就诊(" & Nvl(rsRecord!诊室) & ")"
                End If
            Else
                .TextMatrix(lngRow, 3) = "就诊"
            End If
            
            If .Cols = 5 Then .TextMatrix(lngRow, 4) = Nvl(rsRecord!备注)
            
            lngRow = IIf(mtpPageObj.tpblnShowListHeader, 2, 1) '从第lngRow行开始显示候诊或准备候诊数据
            
            blnExist就诊 = True
        End If
        
        If .Rows < lngRow + 1 Then Exit Sub
        rsClone.Filter = "排队状态=0"
        
        Do While Not rsClone.EOF
            .TextMatrix(lngRow, 0) = " 请"
            .TextMatrix(lngRow, 1) = Format(rsClone!排队号码, "000") & "号"
            .TextMatrix(lngRow, 2) = Nvl(rsClone!患者姓名)
            
            '根据是否有“就诊”确定“候诊”数量
            If mtpPageObj.tpblnShowListHeader Then
                If lngRow >= IIf(blnExist就诊, 2, 1) And lngRow <= IIf(blnExist就诊, 2 + M_LNG_准备就诊显示行 - 1, M_LNG_准备就诊显示行) Then
                    .TextMatrix(lngRow, 3) = "准备就诊"
                Else
                    .TextMatrix(lngRow, 3) = "候诊"
                End If
            Else
                If lngRow >= IIf(blnExist就诊, 1, 0) And lngRow <= IIf(blnExist就诊, M_LNG_准备就诊显示行, M_LNG_准备就诊显示行 - 1) Then
                    .TextMatrix(lngRow, 3) = "准备就诊"
                Else
                    .TextMatrix(lngRow, 3) = "候诊"
                End If
            End If
            
            If .Cols = 5 Then .TextMatrix(lngRow, 4) = Nvl(rsClone!备注)
            
            lngRow = lngRow + 1
            
            If lngRow > mLcdCommonParameter.lngQueueRows - 1 Then Exit Do
            rsClone.MoveNext
        Loop
    End With
Exit Sub
ErrorHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHand
    mlngInterval = 0
    tmrRefreshInterval.Interval = 1000
    
    Call refreshWeekLab
    lblDate.Caption = Format(Now, "yyyy年mm月dd日 hh:mm")
Exit Sub
ErrorHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    Dim i As Integer
    Dim dblHeightScale As Double, dblWidhtScale As Double
    
    '窗体背景
    imgBack.Left = 0
    imgBack.Top = 0
    imgBack.Height = Me.ScaleHeight
    imgBack.Width = Me.ScaleWidth
    
    dblHeightScale = imgBack.Height / mtpPageObj.tpBackImage.lngHeight
    dblWidhtScale = imgBack.Width / mtpPageObj.tpBackImage.lngWidth

    '医院图标
    Call ResizeImg(imgLOGO, dblWidhtScale * mtpPageObj.tpHospitalLOGO.lngLeft, dblHeightScale * mtpPageObj.tpHospitalLOGO.lngTop, dblWidhtScale * mtpPageObj.tpHospitalLOGO.lngWidth, dblHeightScale * mtpPageObj.tpHospitalLOGO.lngHeight)

    '医院名称
    lblHospitalName.Left = dblWidhtScale * mtpPageObj.tpHospitalName.lngLeft
    lblHospitalName.Top = dblHeightScale * mtpPageObj.tpHospitalName.lngTop + dblHeightScale * mtpPageObj.tpHospitalName.lngHeight / 2 - lblHospitalName.Height / 2
    
    '日期
    lblDate.Left = dblWidhtScale * mtpPageObj.tpDate.lngLeft + dblWidhtScale * mtpPageObj.tpDate.lngWidth / 2 - lblDate.Width / 2
    lblDate.Top = dblHeightScale * mtpPageObj.tpDate.lngTop + dblHeightScale * mtpPageObj.tpDate.lngHeight / 2 - lblDate.Height / 2
    
    '星期
    lblWeek.Left = dblWidhtScale * mtpPageObj.tpWeek.lngLeft + dblWidhtScale * mtpPageObj.tpWeek.lngWidth / 2 - lblWeek.Width / 2
    lblWeek.Top = dblHeightScale * mtpPageObj.tpWeek.lngTop + dblHeightScale * mtpPageObj.tpWeek.lngHeight / 2 - lblWeek.Height / 2
    
    '排队列表
    vsfCallingData.Left = dblWidhtScale * mtpPageObj.tpCurQueuedList.lngLeft
    vsfCallingData.Top = dblHeightScale * mtpPageObj.tpCurQueuedList.lngTop
    vsfCallingData.Width = dblWidhtScale * mtpPageObj.tpCurQueuedList.lngWidth
    vsfCallingData.Height = dblHeightScale * mtpPageObj.tpCurQueuedList.lngHeight

    '诊室名称
    lblClinicName.Left = dblWidhtScale * mtpPageObj.tpClinicName.lngLeft + dblWidhtScale * mtpPageObj.tpClinicName.lngWidth / 2 - lblClinicName.Width / 2
    lblClinicName.Top = dblHeightScale * mtpPageObj.tpClinicName.lngTop + dblHeightScale * mtpPageObj.tpClinicName.lngHeight / 2 - lblClinicName.Height / 2
    
    '医生照片
    Call ResizeImg(imgDoctor, dblWidhtScale * mtpPageObj.tpDoctorPhotoAera.lngLeft, dblHeightScale * mtpPageObj.tpDoctorPhotoAera.lngTop, dblWidhtScale * mtpPageObj.tpDoctorPhotoAera.lngWidth, dblHeightScale * mtpPageObj.tpDoctorPhotoAera.lngHeight)

    '医生姓名
    lblDoctorName.Left = dblWidhtScale * mtpPageObj.tpDoctorInfo.lngLeft + dblWidhtScale * mtpPageObj.tpDoctorInfo.lngWidth / 2 - lblDoctorName.Width / 2
    lblDoctorName.Top = dblHeightScale * mtpPageObj.tpDoctorInfo.lngTop + dblHeightScale * mtpPageObj.tpDoctorInfo.lngHeight / 5
    
    '医生职位
    lblDoctorJob.Left = dblWidhtScale * mtpPageObj.tpDoctorInfo.lngLeft + dblWidhtScale * mtpPageObj.tpDoctorInfo.lngWidth / 2 - lblDoctorJob.Width / 2
    lblDoctorJob.Top = dblHeightScale * mtpPageObj.tpDoctorInfo.lngTop + dblHeightScale * mtpPageObj.tpDoctorInfo.lngHeight * 3 / 5
    
    '医生简介
    lblDoctorIntro.Left = dblWidhtScale * mtpPageObj.tpDoctorIntro.lngLeft
    lblDoctorIntro.Top = dblHeightScale * mtpPageObj.tpDoctorIntro.lngTop + 60
    lblDoctorIntro.Width = dblWidhtScale * mtpPageObj.tpDoctorIntro.lngWidth
    lblDoctorIntro.Height = dblHeightScale * mtpPageObj.tpDoctorIntro.lngHeight
    
    '科室简介
    If lblDoctorIntro.Caption = "" Then
        lblDeptInfo.Left = dblWidhtScale * mtpPageObj.tpDoctorIntro.lngLeft
        lblDeptInfo.Top = dblHeightScale * mtpPageObj.tpDoctorPhotoAera.lngTop + 60
        lblDeptInfo.Width = dblWidhtScale * mtpPageObj.tpDoctorIntro.lngWidth
        lblDeptInfo.Height = imgDoctor.Height + lblDoctorIntro.Height
        
        lblDeptInfo.Caption = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "科室简介")
    End If
    
    '呼叫信息
    lblCallContext.Left = imgBack.Width
    lblCallContext.Top = dblHeightScale * mtpPageObj.tpBottomAera.lngTop + dblHeightScale * mtpPageObj.tpBottomAera.lngHeight / 2 - lblRemarkInfo.Height / 2
    
    '备注内容
    lblRemarkInfo.Left = imgBack.Width
    lblRemarkInfo.Top = lblCallContext.Top
    
    With vsfCallingData
        '设置排队列表列宽
        If .Cols = 5 Then
            .ColWidth(0) = .Width * 1 / 9
            .ColWidth(1) = .Width * 2 / 9
            .ColWidth(2) = .Width * 2 / 9
            .ColWidth(3) = .Width * 2 / 9
            .ColWidth(4) = .Width * 2 / 9
        Else
            .ColWidth(0) = .Width * 1 / 7
            .ColWidth(1) = .Width * 2 / 7
            .ColWidth(2) = .Width * 2 / 7
            .ColWidth(3) = .Width * 2 / 7
        End If
        '设置排队列表行高
        For i = 0 To .Rows - 1
            .RowHeight(i) = .Height / .Rows
        Next
    End With
    
    If mLcdCommonParameter.blnFontAutoSizeToList And vsfCallingData.Rows > 1 And vsfCallingData.Cols > 1 Then
        vsfCallingData.Cell(flexcpFontSize, 0, 0, vsfCallingData.Rows - 1, vsfCallingData.Cols - 1) = vsfCallingData.RowHeight(0) / 42
    End If
End Sub

Private Sub tmrRemarkInfo_Timer()
On Error GoTo ErrorHand
    lblRemarkInfo.Left = lblRemarkInfo.Left - 15
    
    If lblRemarkInfo.Left <= -lblRemarkInfo.Width Or lblRemarkInfo.Caption = "" Then
        If lblCallContext.Caption <> "" Then    '滚动显示处于“呼叫中”和“已呼叫”的数据
            lblCallContext.Left = lblCallContext.Left - 15
            
            If lblCallContext.Left <= -lblCallContext.Width Then
                lblCallContext.Left = imgBack.Width
                lblRemarkInfo.Left = imgBack.Width
            End If
        Else
            lblRemarkInfo.Left = imgBack.Width
        End If
    End If
Exit Sub
ErrorHand:
    Debug.Print Err.Description
    Err.Clear
End Sub

Private Sub tmrRefreshInterval_Timer()
    Dim blnExist就诊 As Boolean
    
On Error GoTo ErrorHand
    mlngInterval = mlngInterval + 1

    '当Timer累计的时间小于轮询时间时，无需刷新排队数据
    If mlngInterval < mlngRefreshInterval Then Exit Sub
    '将累计的时间清0
    mlngInterval = 0
    
    Call LoadCallingData(blnExist就诊)
    
    Call SetStyleFont(blnExist就诊)
Exit Sub
ErrorHand:
    Debug.Print Err.Description
    Err.Clear
End Sub

Private Sub tmrTime_Timer()
On Error GoTo ErrorHand
    Call refreshWeekLab
    lblDate.Caption = Format(Now, "yyyy年mm月dd日 hh:mm")
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
