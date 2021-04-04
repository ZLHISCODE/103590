VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmStyle_MultiQueue 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11715
   Icon            =   "frmStyle_MultiQueue.frx":0000
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
   Begin VSFlex8Ctl.VSFlexGrid vsfCallingList 
      Height          =   1215
      Left            =   3240
      TabIndex        =   1
      Top             =   1800
      Width           =   4815
      _cx             =   8493
      _cy             =   2143
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
      ForeColor       =   12582912
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
   Begin VSFlex8Ctl.VSFlexGrid vsfQueueList 
      Height          =   1215
      Left            =   3120
      TabIndex        =   2
      Top             =   3240
      Width           =   4815
      _cx             =   8493
      _cy             =   2143
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
      ForeColor       =   12582912
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
   Begin VB.Image imgLOGO 
      Height          =   720
      Left            =   240
      Picture         =   "frmStyle_MultiQueue.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   840
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
      TabIndex        =   9
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label lblPatientName 
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
      Left            =   480
      TabIndex        =   8
      Top             =   2040
      Width           =   240
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
      TabIndex        =   7
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
Attribute VB_Name = "frmStyle_MultiQueue"
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
Private mstrQueryQueueNames As String
Private mstrQueueListQueryNames As String   '在排队列表上显示的队列名称
Private mstrClinicNames As String       '临床排队业务下的诊室名称

Private mLcdCommonParameter As TLcdCommonParameter

Private Type TPageObj
    tpBackImage  As TRect        '背景(皮肤)
    tpTopArea    As TRect        '顶部
    tpMiddleArea As TRect        '中部
    tpBottomArea As TRect        '底部
    
    tpHospitalLOGO   As TRect    '医院图标
    tpHospitalName   As TRect    '医院名称
    tpWeek           As TRect    '星期
    tpDate           As TRect    '日期
    
    tpCurCallingInf As TRect     '当前呼叫信息
    tpCurCalledList As TRect     '呼叫列表
    tpCurQueuedList As TRect     '准备候诊列表
    
    tplngCallingListMaxRows As Long     '呼叫列表可以显示的总行数
    tplngQueueListMaxRows As Long       '多队列准备就诊列表可以显示的总行数（包括列头）
    tplngQueueListShowNum As Long       '准备就诊列表区域准备就诊列显示数量
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
        .tpTopArea.lngLeft = Val(ReadValue("顶部区域", "左"))
        .tpTopArea.lngTop = Val(ReadValue("顶部区域", "顶"))
        .tpTopArea.lngWidth = Val(ReadValue("顶部区域", "宽"))
        .tpTopArea.lngHeight = Val(ReadValue("顶部区域", "高"))
        
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
        .tpMiddleArea.lngLeft = Val(ReadValue("中部区域", "左"))
        .tpMiddleArea.lngTop = Val(ReadValue("中部区域", "顶"))
        .tpMiddleArea.lngWidth = Val(ReadValue("中部区域", "宽"))
        .tpMiddleArea.lngHeight = Val(ReadValue("中部区域", "高"))
        
        '底部区域
        .tpBottomArea.lngLeft = Val(ReadValue("底部区域", "左"))
        .tpBottomArea.lngTop = Val(ReadValue("底部区域", "顶"))
        .tpBottomArea.lngWidth = Val(ReadValue("底部区域", "宽"))
        .tpBottomArea.lngHeight = Val(ReadValue("底部区域", "高"))
        
        '呼叫信息区域
        .tpCurCallingInf.lngLeft = Val(ReadValue("呼叫信息区域", "左"))
        .tpCurCallingInf.lngTop = Val(ReadValue("呼叫信息区域", "顶"))
        .tpCurCallingInf.lngWidth = Val(ReadValue("呼叫信息区域", "宽"))
        .tpCurCallingInf.lngHeight = Val(ReadValue("呼叫信息区域", "高"))
        
        '呼叫列表区域
        .tpCurCalledList.lngLeft = Val(ReadValue("呼叫列表区域", "左"))
        .tpCurCalledList.lngTop = Val(ReadValue("呼叫列表区域", "顶"))
        .tpCurCalledList.lngWidth = Val(ReadValue("呼叫列表区域", "宽"))
        .tpCurCalledList.lngHeight = Val(ReadValue("呼叫列表区域", "高"))
        
        .tplngCallingListMaxRows = Val(ReadValue("呼叫列表区域", "总行数"))
        
        '准备就诊列表区域
        .tpCurQueuedList.lngLeft = Val(ReadValue("准备就诊列表区域", "左"))
        .tpCurQueuedList.lngTop = Val(ReadValue("准备就诊列表区域", "顶"))
        .tpCurQueuedList.lngWidth = Val(ReadValue("准备就诊列表区域", "宽"))
        .tpCurQueuedList.lngHeight = Val(ReadValue("准备就诊列表区域", "高"))

        .tplngQueueListMaxRows = Val(ReadValue("准备就诊列表区域", "总行数"))
        .tplngQueueListShowNum = Val(ReadValue("准备就诊列表区域", "显示人数"))
    End With
End Sub

Public Sub ISty_RefreshQueueData(Optional ByVal lngQueueId As Long)
'刷新数据
    Call LoadListData(lngQueueId)
    Call SetStyleFont
    
    '数据刷新后将计时器清0
    mlngInterval = 0
End Sub

'打开lcd显示界面
Public Sub ISty_Show(ByVal lngWindowNo As Long)
'lngWindowNo:窗口编号，根据窗口编号读取配置信息，并进行显示
    mlngWindowNo = lngWindowNo
    
    Call InitMonitor    '初始化监视器设置
    
    If Not InitLocalPars Then Exit Sub
    
    Call LoadListData
    
    Call SetStyleFont

    Call Show
End Sub


Public Function ISty_ShowCfg(ByVal lngWindowNo As Long, objOwner As Object) As Boolean
'打开对应的样式配置窗口
    Dim objConfig As frmStyle_CommonCfg
    
    Set objConfig = New frmStyle_CommonCfg
            
    ISty_ShowCfg = objConfig.OpenShowConfig(lngWindowNo, TShowStyle.ssMultiQueue, Me)
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
    Dim i As Integer
    Dim lngCurLCDNo As Long
    Dim strBusinessType As String
    Dim strLCDLocation As String

On Error GoTo ErrorHand
    If gobjFile.FolderExists(App.Path & "\Skin\多队列样式") Then
        mstrStyleTylePath = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "皮肤样式", App.Path & "\Skin\多队列样式\多队列样式深蓝") & ".jpg"
    Else
        mstrStyleTylePath = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "皮肤样式", App.Path & "\zlQueueShow\Skin\多队列样式\多队列样式深蓝") & ".jpg"
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
    mstrQueryQueueNames = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "显示队列", "")
    
    mLcdCommonParameter.blnConvertQueueName = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "转换队列名称", 0)) = 1

    If GetQueueNames(mstrQueryQueueNames) <> "" Then
        If UBound(Split(mLcdCommonParameter.strQueryQueueNames, ",")) + 1 <= 4 Then
            vsfCallingList.Cols = 1
            
            mLcdCommonParameter.lngCallingRows = UBound(Split(mLcdCommonParameter.strQueryQueueNames, ",")) + 1
        Else
            vsfCallingList.Cols = 2
            
            If (UBound(Split(mLcdCommonParameter.strQueryQueueNames, ",")) + 1) Mod 2 = 0 Then
                mLcdCommonParameter.lngCallingRows = (UBound(Split(mLcdCommonParameter.strQueryQueueNames, ",")) + 1) \ 2
            Else
                mLcdCommonParameter.lngCallingRows = (UBound(Split(mLcdCommonParameter.strQueryQueueNames, ",")) + 1) \ 2 + 1
            End If
        End If
        
        mLcdCommonParameter.lngQueueRows = UBound(Split(mLcdCommonParameter.strQueryQueueNames, ",")) + 2
    Else
        mLcdCommonParameter.lngCallingRows = mtpPageObj.tplngCallingListMaxRows
        mLcdCommonParameter.lngQueueRows = mtpPageObj.tplngQueueListMaxRows
        vsfCallingList.Cols = 2
    End If
    
    '根据显示行设置列表背景图
    If vsfCallingList.Cols Mod 2 = 0 Then
        vsfCallingList.WallPaper = CutPicture(mstrStyleTylePath, picTemp, mtpPageObj.tpCurCalledList.lngLeft * Screen.TwipsPerPixelX, mtpPageObj.tpCurCalledList.lngTop * Screen.TwipsPerPixelY, mtpPageObj.tpCurCalledList.lngWidth * Screen.TwipsPerPixelX, mtpPageObj.tpCurCalledList.lngHeight * Screen.TwipsPerPixelY * mLcdCommonParameter.lngCallingRows / mtpPageObj.tplngCallingListMaxRows)
    Else
        vsfCallingList.WallPaper = CutPicture(mstrStyleTylePath, picTemp, mtpPageObj.tpCurCalledList.lngLeft * Screen.TwipsPerPixelX, mtpPageObj.tpCurCalledList.lngTop * Screen.TwipsPerPixelY, (mtpPageObj.tpCurCalledList.lngWidth / 2 - 2) * Screen.TwipsPerPixelX, mtpPageObj.tpCurCalledList.lngHeight * Screen.TwipsPerPixelY * mLcdCommonParameter.lngCallingRows / mtpPageObj.tplngCallingListMaxRows)
    End If

    vsfQueueList.WallPaper = CutPicture(mstrStyleTylePath, picTemp, mtpPageObj.tpCurQueuedList.lngLeft * Screen.TwipsPerPixelX, mtpPageObj.tpCurQueuedList.lngTop * Screen.TwipsPerPixelY, mtpPageObj.tpCurQueuedList.lngWidth * Screen.TwipsPerPixelX, mtpPageObj.tpCurQueuedList.lngHeight * Screen.TwipsPerPixelY * mLcdCommonParameter.lngQueueRows / mtpPageObj.tplngQueueListMaxRows)
    
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
    
    mLcdCommonParameter.blnFontAutoSizeToList = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "列表字体自适应", True)
    
    tmrRefreshInterval.Enabled = True
    tmrRemarkInfo.Enabled = True
    
    InitLocalPars = True
Exit Function
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Function

Private Function GetQueueNames(ByVal strQueryQueueNames As String) As String
'根据排队方式获取队列名称
    Dim i As Integer
    Dim lngPreDeptID As Long
    Dim lngCurDeptID As Long
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    
    mLcdCommonParameter.strQueryQueueNames = ""
    lngPreDeptID = 0
    lngCurDeptID = 0

    If mLcdCommonParameter.blnConvertQueueName Then    '转换成老版本格式的队列名称
        For i = 0 To UBound(Split(strQueryQueueNames, ","))
            lngCurDeptID = Split(Split(Split(strQueryQueueNames, ",")(i), "|")(1), "_")(0)
    
            Select Case glngBusinessType
                Case TBusinessType.btClinical   '队列名称存储规则，"科室名ID1,科室ID2,科室ID3"； 如"63,64,65"
                    mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & lngCurDeptID
                    
                    If InStr(Split(mstrQueryQueueNames, ",")(i), "科室队列") > 0 Then
                        mstrClinicNames = mstrClinicNames & "," & lngCurDeptID & "-"
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), "|")(0)
                    Else
                        mstrClinicNames = mstrClinicNames & "," & lngCurDeptID & "-" & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                    End If
                    
                Case TBusinessType.btPacs       '队列名称存储规则
                    If InStr(Split(strQueryQueueNames, ",")(i), "科室队列") > 0 Then     '按科室排队，"编码1-科室名1,编码2-科室名2,编码3-科室名3"； 如"050202-放射科,050203-CT检查室"
                        mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & Split(Split(Split(Split(strQueryQueueNames, ",")(i), "|")(1), ":")(0), "_")(1) & "-" & Split(Split(strQueryQueueNames, ",")(i), "|")(0)
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), "|")(0)
                    Else                                            '按执行间排队
                        mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & lngCurDeptID & ":" & Split(Split(Split(strQueryQueueNames, ",")(i), "|")(1), ":")(1)
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                    End If
                    
                Case TBusinessType.btPeis       '队列名称存储规则，"站点名:执行将"；如"站点一:执行间1,站点二:执行间1"
                    '根据执行间的科室ID找到对应的站点名称
                    strSql = "select 站点名称 from 体检站点分布 where 执行科室id=[1]"
                    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取站点名称", lngCurDeptID)
                    
                    If rsRecord.RecordCount > 0 Then
                        mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & _
                                                                 Nvl(rsRecord!站点名称) & ":" & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                    End If
                'case
                '
                '
            End Select
            
            lngPreDeptID = lngCurDeptID
        Next

    Else            '''''''''新版队列名称格式
        For i = 0 To UBound(Split(mstrQueryQueueNames, ","))
            lngCurDeptID = Split(Split(Split(strQueryQueueNames, ",")(i), "|")(1), "_")(0)
            
            Select Case glngBusinessType
                Case TBusinessType.btClinical   '队列名称存储规则，"科室ID1,科室ID2,科室ID3"； 如"63,64,65"
                    mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & lngCurDeptID
                    
                    If InStr(Split(mstrQueryQueueNames, ",")(i), "科室队列") > 0 Then
                        mstrClinicNames = mstrClinicNames & "," & lngCurDeptID & "-"
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), "|")(0)
                    Else
                        mstrClinicNames = mstrClinicNames & "," & lngCurDeptID & "-" & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                    End If
                    
                Case TBusinessType.btPacs       '队列名称存储规则，"科室名-执行间1,科室名-执行间2,科室名-执行间3"； 如"放射科-CT一检查室,放射科-CT二检查室,放射科-CT三检查室"
                    mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & _
                                                             Split(Split(strQueryQueueNames, ",")(i), "|")(0) & "-" & Split(Split(Split(strQueryQueueNames, ",")(i), "|")(1), ":")(1)
                    
                    If InStr(Split(mstrQueryQueueNames, ",")(i), "未分配队列") > 0 Then
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), "|")(0)
                    Else
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                    End If
                Case TBusinessType.btPeis       '队列名称存储规则，"站点名:执行将"；如"站点一:执行间1,站点二:执行间1"
                    '根据执行间的科室ID找到对应的站点名称
                    strSql = "select 站点名称 from 体检站点分布 where 执行科室id=[1]"
                    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取站点名称", lngCurDeptID)
                    
                    If rsRecord.RecordCount > 0 Then
                        mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & _
                                                                 Nvl(rsRecord!站点名称) & ":" & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                    End If
                'case
                '
                '
            End Select
            
            lngPreDeptID = lngCurDeptID
        Next
    End If
    mLcdCommonParameter.strQueryQueueNames = Mid(mLcdCommonParameter.strQueryQueueNames, 2)
    mstrClinicNames = Mid(mstrClinicNames, 2)
    mstrQueueListQueryNames = Mid(mstrQueueListQueryNames, 2)
    GetQueueNames = mLcdCommonParameter.strQueryQueueNames
End Function

Private Sub InitDataList()
'初始化数据列表
    Dim i As Integer
    
    With vsfCallingList
        .Rows = mLcdCommonParameter.lngCallingRows
        
        For i = 0 To .Rows - 1
            .RowHeight(i) = .Height / .Rows
        Next
        
        For i = 0 To .Cols - 1
            .ColWidth(i) = .Width / .Cols
        Next
        
        If .Rows > 0 And .Cols > 0 Then .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
    End With
    
    With vsfQueueList
        .Cols = 3
        .Rows = mLcdCommonParameter.lngQueueRows
        
        For i = 0 To .Rows - 1
            .RowHeight(i) = .Height / .Rows
        Next
        
        .ColWidth(0) = .Width * 2 / 10
        .ColWidth(1) = .Width * 6 / 10
        .ColWidth(2) = .Width * 2 / 10
        
        .TextMatrix(0, 0) = "  队列名称"
        .TextMatrix(0, 1) = "准备就诊列表"
        .TextMatrix(0, 2) = "排队人数"
        
        If .Rows > 0 And .Cols > 0 Then
            .Cell(flexcpAlignment, 0, 0) = flexAlignLeftCenter
            .Cell(flexcpAlignment, 0, 1) = flexAlignCenterCenter
        End If
        
        If .Rows > 1 And .Cols > 0 Then
            .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
            .Cell(flexcpAlignment, 0, .Cols - 1, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        End If
        
        If mstrQueryQueueNames = "" Then Exit Sub
        
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = "  " & Split(mstrQueueListQueryNames, ",")(i - 1)
            .TextMatrix(i, 2) = "共 0  人"
        Next
    End With
End Sub

Private Sub LoadListData(Optional ByVal lngQueueId As Long)
'加载排队呼叫列表数据
    Dim i As Integer, j As Integer, k As Integer
    Dim lngRow As Long, lngCol As Long
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    Dim rsClone As ADODB.Recordset
    Dim strSortStyle As String      '排序方式
    Dim strQueuePatients As String
    Dim dblHeightScale As Double, dblWidhtScale As Double
    Dim strTemp As String

On Error GoTo ErrorHand:

    Call InitDataList
    
    lblPatientName.Caption = ""
    lblClinicName.Caption = ""
    
    If mLcdCommonParameter.strQueryQueueNames = "" Or glngBusinessType < 0 Then Exit Sub
    
    strSql = "select Zl_排队叫号队列_获取排序方式([1]) as 排序方式 from dual"
    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取排序方式", glngBusinessType)
    
    If rsRecord.RecordCount > 0 Then strSortStyle = Nvl(rsRecord!排序方式)
    
    Select Case glngBusinessType
        Case TBusinessType.btClinical
            strSql = "select distinct A.ID,A.排队号码,A.患者姓名,A.排队状态,A.呼叫时间,A.备注,A.排队序号,A.队列名称,A.诊室,C.名称 " & _
                     "from 排队叫号队列 A, " & _
                     "(select 队列名称,诊室 from " & _
                     "(select Column_Value as 队列名称 from Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) m) C, " & _
                     "(select substr(Column_Value,1,instr(Column_Value,'-')-1) as 科室,substr(Column_Value,instr(Column_Value,'-')+1) as 诊室 from Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) n) D " & _
                     "where C.队列名称=D.科室) B,部门表 C " & _
                     "Where A.队列名称 =B.队列名称 and a.科室id=c.id and (A.诊室=B.诊室 or A.诊室 is null  or B.诊室 is null) and 业务类型=[3] and 排队状态 in (0,1,7,8,9) And 排队时间 > trunc(sysdate)" & _
                     IIf(mLcdCommonParameter.strFilter <> "", " and " & mLcdCommonParameter.strFilter, "") & " " & _
                     IIf(strSortStyle <> "", " order by " & strSortStyle, "")
                     
            Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取排队信息", mLcdCommonParameter.strQueryQueueNames, mstrClinicNames, glngBusinessType)
            
        Case TBusinessType.btPacs
            strSql = "select ID,排队号码,患者姓名,排队状态,呼叫时间,备注,排队序号,队列名称,诊室 from 排队叫号队列 A,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) B " & _
                     "Where A.队列名称 =B.Column_Value and 业务类型=[2] and 排队状态 in (0,1,5,6,7,8,9) And 排队时间 > trunc(sysdate) " & _
                     IIf(mLcdCommonParameter.strFilter <> "", " and " & mLcdCommonParameter.strFilter, "") & " " & _
                     IIf(strSortStyle <> "", " order by " & strSortStyle, "")
                     
            Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取排队信息", mLcdCommonParameter.strQueryQueueNames, glngBusinessType)
            
        Case TBusinessType.btPeis
            strSql = "select ID,排队号码,患者姓名,排队状态,呼叫时间,备注,排队序号,队列名称,诊室 from 排队叫号队列 A,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) B " & _
                     "Where A.队列名称 =B.Column_Value and 业务类型=[2] and 排队状态 in (0,1,7,8,9) And 排队时间 > trunc(sysdate) " & _
                     IIf(mLcdCommonParameter.strFilter <> "", " and " & mLcdCommonParameter.strFilter, "") & " " & _
                     IIf(strSortStyle <> "", " order by " & strSortStyle, "")
                     
            Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取排队信息", mLcdCommonParameter.strQueryQueueNames, glngBusinessType)
        'case
        '
        '
    End Select
    If rsRecord.RecordCount <= 0 Then Exit Sub
    Set rsClone = rsRecord.Clone

    '加载正在呼叫的科室及病人信息
    If gstrCompareVersion < "010.034.000" Then
        rsRecord.Filter = "id=" & lngQueueId
    Else
        rsRecord.Filter = "排队状态=9"
        If rsRecord.RecordCount <= 0 Then rsRecord.Filter = "排队状态=1"
    End If
    
    '当前呼叫信息
    If rsRecord.RecordCount > 0 Then
        lblPatientName.Caption = Format(Nvl(rsRecord!排队号码), "000") & "号 " & Nvl(rsRecord!患者姓名)
        If glngBusinessType = TBusinessType.btClinical Then
            lblClinicName.Caption = IIf(Nvl(rsRecord!诊室) = "", Nvl(rsRecord!名称), Nvl(rsRecord!诊室))
        Else
            lblClinicName.Caption = Nvl(rsRecord!诊室)
        End If
        
        '重新设置当前呼叫信息显示位置
        dblHeightScale = imgBack.Height / mtpPageObj.tpBackImage.lngHeight
        dblWidhtScale = imgBack.Width / mtpPageObj.tpBackImage.lngWidth
        
        lblPatientName.Left = dblWidhtScale * mtpPageObj.tpCurCallingInf.lngLeft + (dblWidhtScale * mtpPageObj.tpCurCallingInf.lngWidth - lblPatientName.Width) / 2
        lblPatientName.Top = dblHeightScale * mtpPageObj.tpCurCallingInf.lngTop + 60
        
        lblClinicName.Left = dblWidhtScale * mtpPageObj.tpCurCallingInf.lngLeft + (dblWidhtScale * mtpPageObj.tpCurCallingInf.lngWidth - lblClinicName.Width) / 2
        lblClinicName.Top = lblPatientName.Top + 780
    End If
    
    For i = 1 To vsfQueueList.Rows - 1
        '加载准备就诊列表数据
        rsRecord.Filter = ""
        strQueuePatients = ""
        
        If mstrClinicNames <> "" Then strTemp = Split(Split(mstrClinicNames, ",")(i - 1), "-")(1)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''排队状态=5''''''''''''''''''''''''''''''''''''''''''''''''''''加载候诊数据
        If strTemp = "" Then
            rsRecord.Filter = "队列名称='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And 排队状态=5"
        Else
            rsRecord.Filter = "队列名称='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And 诊室='" & strTemp & "' And 排队状态=0"
        End If
        
        For j = 0 To rsRecord.RecordCount - 1
            strQueuePatients = strQueuePatients & ", " & Format(Nvl(rsRecord!排队号码), "000") & "-" & Nvl(rsRecord!患者姓名) & IIf(Nvl(rsRecord!备注) <> "", "(" & Nvl(rsRecord!备注) & ")", "")
            If j = mtpPageObj.tplngQueueListShowNum - 1 Then Exit For '每个队列显示准备就诊数据
            rsRecord.MoveNext
        Next
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''排队状态=6
        If strTemp = "" Then
            rsRecord.Filter = "队列名称='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And 排队状态=6"
        Else
            rsRecord.Filter = "队列名称='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And 诊室='" & strTemp & "' And 排队状态=0"
        End If
        
        For j = 0 To rsRecord.RecordCount - 1
            strQueuePatients = strQueuePatients & ", " & Format(Nvl(rsRecord!排队号码), "000") & "-" & Nvl(rsRecord!患者姓名) & IIf(Nvl(rsRecord!备注) <> "", "(" & Nvl(rsRecord!备注) & ")", "")
            If j = mtpPageObj.tplngQueueListShowNum - 1 Then Exit For '每个队列显示准备就诊数据
            rsRecord.MoveNext
        Next
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''排队状态=0
        If strTemp = "" Then
            rsRecord.Filter = "队列名称='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And 排队状态=0 "
        End If

        For j = 0 To rsRecord.RecordCount - 1
            strQueuePatients = strQueuePatients & ", " & Format(Nvl(rsRecord!排队号码), "000") & "-" & Nvl(rsRecord!患者姓名) & IIf(Nvl(rsRecord!备注) <> "", "(" & Nvl(rsRecord!备注) & ")", "")
            If j = mtpPageObj.tplngQueueListShowNum - 1 Then Exit For '每个队列显示准备就诊数据
            rsRecord.MoveNext
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''加载候诊数据完成
        vsfQueueList.TextMatrix(i, 1) = Mid(strQueuePatients, 2)
        vsfQueueList.TextMatrix(i, 2) = "共 " & strFormat(rsRecord.RecordCount) & "人"
        
        '加载呼叫列表数据
        If strTemp = "" Then
            rsClone.Filter = "队列名称='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And 排队状态=9"
            If rsClone.RecordCount <= 0 Then rsClone.Filter = "队列名称='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And 排队状态=1"
            If rsClone.RecordCount <= 0 Then rsClone.Filter = "队列名称='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And 排队状态=7"
            If rsClone.RecordCount <= 0 Then rsClone.Filter = "队列名称='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And 排队状态=8"
        Else
            rsClone.Filter = "队列名称='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And 诊室='" & strTemp & "' And 排队状态=9"
            If rsClone.RecordCount <= 0 Then rsClone.Filter = "队列名称='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And 诊室='" & strTemp & "' And 排队状态=1"
            If rsClone.RecordCount <= 0 Then rsClone.Filter = "队列名称='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And 诊室='" & strTemp & "' And 排队状态=7"
            If rsClone.RecordCount <= 0 Then rsClone.Filter = "队列名称='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And 诊室='" & strTemp & "' And 排队状态=8"
        End If
        
        If rsClone.RecordCount > 0 Then
            rsClone.Sort = "呼叫时间 desc"
            k = k + 1
            
            If vsfCallingList.Cols = 1 Then
                If glngBusinessType = TBusinessType.btClinical Then
                    vsfCallingList.TextMatrix(k - 1, 0) = "●请 " & Format(Nvl(rsClone!排队号码), "000") & " 号 " & Nvl(rsClone!患者姓名) & " 到 " & IIf(Nvl(rsClone!诊室) = "", Nvl(rsClone!名称), Nvl(rsClone!诊室)) & " 就诊 "
                Else
                    vsfCallingList.TextMatrix(k - 1, 0) = "●请 " & Format(Nvl(rsClone!排队号码), "000") & " 号 " & Nvl(rsClone!患者姓名) & " 到 " & Nvl(rsClone!诊室) & " 就诊 "
                End If
            Else
                If glngBusinessType = TBusinessType.btClinical Then
                    vsfCallingList.TextMatrix(lngRow, lngCol) = "●请 " & Format(Nvl(rsClone!排队号码), "000") & " 号 " & Nvl(rsClone!患者姓名) & " 到 " & IIf(Nvl(rsClone!诊室) = "", Nvl(rsClone!名称), Nvl(rsClone!诊室)) & " 就诊 "
                Else
                    vsfCallingList.TextMatrix(lngRow, lngCol) = "●请 " & Format(Nvl(rsClone!排队号码), "000") & " 号 " & Nvl(rsClone!患者姓名) & " 到 " & Nvl(rsClone!诊室) & " 就诊 "
                End If
                lngCol = k Mod 2
                lngRow = k \ 2
            End If
        End If
    Next
    
    lblCallContext.Caption = ""
    
    '获取处于“呼叫中”和“已呼叫”的数据
    If mLcdCommonParameter.blnScrollDisplay Then
        For i = 1 To vsfQueueList.Rows - 1
            If mstrClinicNames <> "" Then strTemp = Split(Split(mstrClinicNames, ",")(i - 1), "-")(1)
            
            If strTemp = "" Then
                rsRecord.Filter = "队列名称='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And 排队状态=7"
                rsClone.Filter = "队列名称='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And 排队状态=1"
            Else
                rsRecord.Filter = "队列名称='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And 诊室='" & strTemp & "' And 排队状态=7"
                rsClone.Filter = "队列名称='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And 诊室='" & strTemp & "' And 排队状态=1"
            End If
            
            If rsRecord.RecordCount > 0 Then
                rsRecord.Sort = "呼叫时间 asc"
                rsRecord.MoveFirst
                
                For j = 0 To IIf(rsClone.RecordCount > 0, rsRecord.RecordCount - 1, rsRecord.RecordCount - 2)
                    If glngBusinessType = TBusinessType.btClinical Then
                        lblCallContext.Caption = lblCallContext.Caption & " ●" & Format(Nvl(rsRecord!排队号码), "000") & "号 " & Nvl(rsRecord!患者姓名) & " 到 " & IIf(Nvl(rsRecord!诊室) = "", Nvl(rsRecord!名称), Nvl(rsRecord!诊室)) & " 就诊"
                    Else
                        lblCallContext.Caption = lblCallContext.Caption & " ●" & Format(Nvl(rsRecord!排队号码), "000") & "号 " & Nvl(rsRecord!患者姓名) & " 到 " & Nvl(rsRecord!诊室) & " 就诊"
                    End If
                    
                    rsRecord.MoveNext
                Next
            End If
        Next
        
        lblRemarkInfo.Caption = ""
    End If
    
    If lblCallContext.Caption = "" Then lblCallContext.Caption = "请未叫到号的患者耐心等待！"
    
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Function strFormat(ByVal strVal As String) As String
    Dim lngLength As Long
    
    lngLength = Len(strVal)
    
    strFormat = strVal & String(3 - Len(strVal), " ")
End Function

Private Sub SetStyleFont()
'设置界面各控件字体属性
    Dim i As Integer
    Dim strFontPropertys As String           '格式:"字体:宋体|字号:20|粗体:FALSE|斜体:FALSE|下划线:FALSE|前景色:16777215"
    Dim strFontPropertys1 As String
    Dim strFontPropertys2 As String
    Dim strFontPropertys3 As String
    Dim strFontProperty() As String
    Dim strFontProperty1() As String
    Dim strFontProperty2() As String
    Dim strFontProperty3() As String
    
    Dim strRegPath As String
On Error GoTo ErrorHand
    strRegPath = G_STR_REGPATH & "\多队列样式"
    
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
    
    '呼叫信息
    strFontPropertys = Trim(ReadValue("字体设置", "呼叫信息字体", "字体:宋体|字号:26|粗体:TRUE|斜体:FALSE|下划线:FALSE|前景色:194300"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblPatientName, strFontProperty)
        Call SetControlFont(lblClinicName, strFontProperty)
    End If
    
    '备注内容
    strFontPropertys = Trim(ReadValue("字体设置", "备注内容字体", "字体:宋体|字号:26|粗体:TRUE|斜体:FALSE|下划线:FALSE|前景色:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblRemarkInfo, strFontProperty)
        Call SetControlFont(lblCallContext, strFontProperty)
    End If
    
    '列表字体
    strFontPropertys1 = Trim(ReadValue("字体设置", "就诊状态行字体", "字体:宋体|字号:18|粗体:TRUE|前景色:55871"))
    strFontPropertys2 = Trim(ReadValue("字体设置", "排队列表标题字体", "字体:宋体|字号:20|粗体:TRUE|前景色:14721613"))
    strFontPropertys3 = Trim(ReadValue("字体设置", "准备就诊状态行字体", "字体:宋体|字号:18|粗体:FALSE|前景色:16777215"))
    
    strFontProperty1 = Split(strFontPropertys1, "|")
    strFontProperty2 = Split(strFontPropertys2, "|")
    strFontProperty3 = Split(strFontPropertys3, "|")
    
    For i = 0 To vsfCallingList.Rows - 1
        SetVSFListFont vsfCallingList, i, strFontProperty1
    Next
    
    SetVSFListFont vsfQueueList, 0, strFontProperty2
    
    For i = 1 To vsfQueueList.Rows - 1
        SetVSFListFont vsfQueueList, i, strFontProperty3
    Next
    
    If mLcdCommonParameter.blnFontAutoSizeToList Then
        If vsfCallingList.Rows >= 1 And vsfCallingList.Cols >= 1 Then
            If vsfCallingList.Cols = 1 Then
                vsfCallingList.Cell(flexcpFontSize, 0, 0, vsfCallingList.Rows - 1, vsfCallingList.Cols - 1) = vsfCallingList.RowHeight(0) / 34
            Else
                vsfCallingList.Cell(flexcpFontSize, 0, 0, vsfCallingList.Rows - 1, vsfCallingList.Cols - 1) = vsfCallingList.RowHeight(0) / 46
            End If
        End If
        
        If vsfQueueList.Rows >= 1 And vsfQueueList.Cols > 1 Then
            vsfQueueList.Cell(flexcpFontSize, 0, 0, vsfQueueList.Rows - 1, vsfQueueList.Cols - 1) = vsfQueueList.RowHeight(0) / 42
        End If
    End If
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
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
    
    '当前呼叫信息
    lblPatientName.Left = dblWidhtScale * mtpPageObj.tpCurCallingInf.lngLeft + (dblWidhtScale * mtpPageObj.tpCurCallingInf.lngWidth - lblPatientName.Width) / 2
    lblPatientName.Top = dblHeightScale * mtpPageObj.tpCurCallingInf.lngTop + 60
    
    lblClinicName.Left = dblWidhtScale * mtpPageObj.tpCurCallingInf.lngLeft + (dblWidhtScale * mtpPageObj.tpCurCallingInf.lngWidth - lblClinicName.Width) / 2
    lblClinicName.Top = lblPatientName.Top + 780

    '呼叫列表
    vsfCallingList.Left = dblWidhtScale * mtpPageObj.tpCurCalledList.lngLeft
    vsfCallingList.Top = dblHeightScale * mtpPageObj.tpCurCalledList.lngTop
    vsfCallingList.Height = dblHeightScale * mtpPageObj.tpCurCalledList.lngHeight
    vsfCallingList.Width = dblWidhtScale * mtpPageObj.tpCurCalledList.lngWidth
    
    '准备就诊列表
    vsfQueueList.Left = dblWidhtScale * mtpPageObj.tpCurQueuedList.lngLeft
    vsfQueueList.Top = dblHeightScale * mtpPageObj.tpCurQueuedList.lngTop
    vsfQueueList.Height = dblHeightScale * mtpPageObj.tpCurQueuedList.lngHeight
    vsfQueueList.Width = dblWidhtScale * mtpPageObj.tpCurQueuedList.lngWidth
    
    '呼叫信息
    lblCallContext.Left = imgBack.Width
    lblCallContext.Top = dblHeightScale * mtpPageObj.tpBottomArea.lngTop + dblHeightScale * mtpPageObj.tpBottomArea.lngHeight / 2 - lblRemarkInfo.Height / 2
    
    '备注内容
    lblRemarkInfo.Left = imgBack.Width
    lblRemarkInfo.Top = lblCallContext.Top
    
    For i = 0 To vsfCallingList.Rows - 1
        vsfCallingList.RowHeight(i) = vsfCallingList.Height / vsfCallingList.Rows
    Next
        
    For i = 0 To vsfQueueList.Rows - 1
        vsfQueueList.RowHeight(i) = vsfQueueList.Height / vsfQueueList.Rows
    Next
    
    If mLcdCommonParameter.blnFontAutoSizeToList Then
        If vsfCallingList.Rows >= 1 And vsfCallingList.Cols >= 1 Then
            If vsfCallingList.Cols = 1 Then
                vsfCallingList.Cell(flexcpFontSize, 0, 0, vsfCallingList.Rows - 1, vsfCallingList.Cols - 1) = vsfCallingList.RowHeight(0) / 34
            Else
                vsfCallingList.Cell(flexcpFontSize, 0, 0, vsfCallingList.Rows - 1, vsfCallingList.Cols - 1) = vsfCallingList.RowHeight(0) / 46
            End If
        End If
        
        If vsfQueueList.Rows >= 1 And vsfQueueList.Cols > 1 Then
            vsfQueueList.Cell(flexcpFontSize, 0, 0, vsfQueueList.Rows - 1, vsfQueueList.Cols - 1) = vsfQueueList.RowHeight(0) / 42
        End If
    End If
End Sub


Private Sub tmrRefreshInterval_Timer()
On Error GoTo ErrorHand
    mlngInterval = mlngInterval + 1

    '当Timer累计的时间小于轮询时间时，无需刷新排队数据
    If mlngInterval < mlngRefreshInterval Then Exit Sub
    '将累计的时间清0
    mlngInterval = 0
    
    Call LoadListData
    
    Call SetStyleFont
Exit Sub
ErrorHand:
    Debug.Print Err.Description
    Err.Clear
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

Private Sub tmrTime_Timer()
On Error GoTo ErrorHand
    Call refreshWeekLab
    lblDate.Caption = Format(Now, "yyyy年mm月dd日 hh:mm")
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
