VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPersonPhoto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "照片采集"
   ClientHeight    =   5355
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8625
   Icon            =   "frmPersonPhoto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "照片"
      Height          =   2355
      Left            =   5025
      TabIndex        =   2
      Top             =   45
      Width           =   3555
      Begin VB.CommandButton cmdLoad 
         Caption         =   "摄像加载(&L)"
         Height          =   350
         Left            =   2025
         TabIndex        =   11
         Top             =   1830
         Width           =   1365
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "文件加载(&F)"
         Height          =   350
         Left            =   2025
         TabIndex        =   10
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "清除照片(&C)"
         Height          =   350
         Left            =   2025
         TabIndex        =   9
         Top             =   720
         Width           =   1365
      End
      Begin VB.PictureBox picPhoto 
         AutoRedraw      =   -1  'True
         Height          =   1984
         Left            =   165
         ScaleHeight     =   1920
         ScaleWidth      =   1350
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   1417
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "摄像"
      Height          =   2760
      Left            =   5010
      TabIndex        =   1
      Top             =   2550
      Width           =   3570
      Begin VB.CommandButton cmdSource 
         Caption         =   "来源调整(&S)"
         Height          =   350
         Left            =   2055
         TabIndex        =   8
         Top             =   1065
         Width           =   1365
      End
      Begin VB.CommandButton cmdStyle 
         Caption         =   "格式调整(&A)"
         Height          =   350
         Left            =   2055
         TabIndex        =   7
         Top             =   645
         Width           =   1365
      End
      Begin VB.ComboBox cboDev 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   195
         Width           =   2220
      End
      Begin VB.PictureBox picFilm 
         Height          =   1984
         Left            =   210
         ScaleHeight     =   1920
         ScaleWidth      =   1350
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   645
         Width           =   1417
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "采集设备(&D)"
         Height          =   180
         Left            =   195
         TabIndex        =   6
         Top             =   255
         Width           =   990
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   5265
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   4935
      _cx             =   8705
      _cy             =   9287
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12698049
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   -1  'True
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
      ComboSearch     =   3
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
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   435
         Y2              =   1650
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   5325
      Top             =   5775
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":000C
            Key             =   "公共"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":03A6
            Key             =   "报告"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":063C
            Key             =   "状态"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":09D6
            Key             =   "住院"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":0D70
            Key             =   "单据"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":110A
            Key             =   "附加"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":14A4
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":173A
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":19D0
            Key             =   "GChecked"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":1C66
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":1EFC
            Key             =   "Checked"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   6300
      Top             =   5715
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPersonPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean
Private mblnStarted As Boolean
Private mlng病人id As Long

Private mbytPopMenu As Byte
'--------------------------------------------------------
'功  能：连接视频输入设备。
'编制人：曾超
'编制日期：2005.11.8
'过程函数清单：
'       mConnCapDevice() 创建设备连接
'       mGetCapSureDevice()
'       mParentWindowResize
'修改记录：
'
'-------------------------------------------------------
Private Const WM_USER As Long = &H400
Private Const WM_CAP_START As Long = WM_USER

Private Const WM_CAP_GET_CAPSTREAMPTR As Long = WM_CAP_START + 1

Private Const WM_CAP_SET_CALLBACK_ERROR As Long = WM_CAP_START + 2
Private Const WM_CAP_SET_CALLBACK_STATUS As Long = WM_CAP_START + 3
Private Const WM_CAP_SET_CALLBACK_YIELD As Long = WM_CAP_START + 4
Private Const WM_CAP_SET_CALLBACK_FRAME As Long = WM_CAP_START + 5
Private Const WM_CAP_SET_CALLBACK_VIDEOSTREAM As Long = WM_CAP_START + 6
Private Const WM_CAP_SET_CALLBACK_WAVESTREAM As Long = WM_CAP_START + 7
Private Const WM_CAP_GET_USER_DATA As Long = WM_CAP_START + 8
Private Const WM_CAP_SET_USER_DATA As Long = WM_CAP_START + 9
    
Private Const WM_CAP_DRIVER_CONNECT As Long = WM_CAP_START + 10
Private Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP_START + 11
Private Const WM_CAP_DRIVER_GET_NAME As Long = WM_CAP_START + 12
Private Const WM_CAP_DRIVER_GET_VERSION As Long = WM_CAP_START + 13
Private Const WM_CAP_DRIVER_GET_CAPS As Long = WM_CAP_START + 14

Private Const WM_CAP_FILE_SET_CAPTURE_FILE As Long = WM_CAP_START + 20
Private Const WM_CAP_FILE_GET_CAPTURE_FILE As Long = WM_CAP_START + 21
Private Const WM_CAP_FILE_ALLOCATE As Long = WM_CAP_START + 22
Private Const WM_CAP_FILE_SAVEAS As Long = WM_CAP_START + 23
Private Const WM_CAP_FILE_SET_INFOCHUNK As Long = WM_CAP_START + 24
Private Const WM_CAP_FILE_SAVEDIB As Long = WM_CAP_START + 25

Private Const WM_CAP_EDIT_COPY As Long = WM_CAP_START + 30

Private Const WM_CAP_SET_AUDIOFORMAT As Long = WM_CAP_START + 35
Private Const WM_CAP_GET_AUDIOFORMAT As Long = WM_CAP_START + 36

Private Const WM_CAP_DLG_VIDEOFORMAT As Long = WM_CAP_START + 41
Private Const WM_CAP_DLG_VIDEOSOURCE As Long = WM_CAP_START + 42
Private Const WM_CAP_DLG_VIDEODISPLAY As Long = WM_CAP_START + 43
Private Const WM_CAP_GET_VIDEOFORMAT As Long = WM_CAP_START + 44
Private Const WM_CAP_SET_VIDEOFORMAT As Long = WM_CAP_START + 45
Private Const WM_CAP_DLG_VIDEOCOMPRESSION As Long = WM_CAP_START + 46

Private Const WM_CAP_SET_PREVIEW As Long = WM_CAP_START + 50
Private Const WM_CAP_SET_OVERLAY As Long = WM_CAP_START + 51
Private Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP_START + 52
Private Const WM_CAP_SET_SCALE As Long = WM_CAP_START + 53
Private Const WM_CAP_GET_STATUS As Long = WM_CAP_START + 54
Private Const WM_CAP_SET_SCROLL As Long = WM_CAP_START + 55

Private Const WM_CAP_GRAB_FRAME As Long = WM_CAP_START + 60
Private Const WM_CAP_GRAB_FRAME_NOSTOP As Long = WM_CAP_START + 61

Private Const WM_CAP_SEQUENCE As Long = WM_CAP_START + 62
Private Const WM_CAP_SEQUENCE_NOFILE As Long = WM_CAP_START + 63
Private Const WM_CAP_SET_SEQUENCE_SETUP As Long = WM_CAP_START + 64
Private Const WM_CAP_GET_SEQUENCE_SETUP As Long = WM_CAP_START + 65
Private Const WM_CAP_SET_MCI_DEVICE As Long = WM_CAP_START + 66
Private Const WM_CAP_GET_MCI_DEVICE As Long = WM_CAP_START + 67
Private Const WM_CAP_STOP As Long = WM_CAP_START + 68
Private Const WM_CAP_ABORT As Long = WM_CAP_START + 69

Private Const WM_CAP_SINGLE_FRAME_OPEN As Long = WM_CAP_START + 70
Private Const WM_CAP_SINGLE_FRAME_CLOSE As Long = WM_CAP_START + 71
Private Const WM_CAP_SINGLE_FRAME As Long = WM_CAP_START + 72

Private Const WM_CAP_PAL_OPEN As Long = WM_CAP_START + 80
Private Const WM_CAP_PAL_SAVE As Long = WM_CAP_START + 81
Private Const WM_CAP_PAL_PASTE As Long = WM_CAP_START + 82
Private Const WM_CAP_PAL_AUTOCREATE As Long = WM_CAP_START + 83
Private Const WM_CAP_PAL_MANUALCREATE As Long = WM_CAP_START + 84

Private Const WM_CAP_SET_CALLBACK_CAPCONTROL As Long = WM_CAP_START + 85

Private Const WS_CHILD As Long = &H40000000
Private Const WS_VISIBLE As Long = &H10000000
Private Const SWP_NOSIZE As Long = &H1&
Private Const SWP_NOMOVE As Long = &H2&
Private Const SWP_NOZORDER As Long = &H4&
Private Const SWP_NOSENDCHANGING As Long = &H400&
Private Const HWND_BOTTOM As Long = 1&

Private hCapWnd As Long                          '采集窗体句柄

Private Type VFWPOINT
        x As Long
        y As Long
End Type

Private Type CAPSTATUS
    uiImageWidth As Long
    uiImageHeight As Long
    fLiveWindow As Long
    fOverlayWindow As Long
    fScale As Long
    ptScroll As VFWPOINT
    fUsingDefaultPalette As Long
    fAudioHardware As Long
    fCapFileExists As Long
    dwCurrentVideoFrame As Long
    dwCurrentVideoFramesDropped As Long
    dwCurrentWaveSamples As Long
    dwCurrentTimeElapsedMS As Long
    hPalCurrent As Long
    fCapturingNow As Long
    dwReturn As Long
    wNumVideoAllocated As Long
    wNumAudioAllocated As Long
End Type



'得到采集驱动列表
Private Declare Function capGetDriverDescription Lib "avicap32.dll" Alias "capGetDriverDescriptionA" _
                                        (ByVal dwDriverIndex As Long, _
                                        ByVal lpszName As String, _
                                        ByVal cbName As Long, _
                                        ByVal lpszVer As String, _
                                        ByVal cbVer As Long) As Long
'创建采集窗口
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" _
                                        (ByVal lpszWindowName As String, _
                                        ByVal dwStyle As Long, _
                                        ByVal x As Long, _
                                        ByVal y As Long, _
                                        ByVal nWidth As Long, _
                                        ByVal nHeight As Long, _
                                        ByVal hwndParent As Long, _
                                        ByVal nID As Long) As Long
'消息发送
Private Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" _
                                            (ByVal hWnd As Long, _
                                            ByVal wMsg As Long, _
                                            ByVal wParam As Long, _
                                            ByVal lParam As Long) As Long
Private Declare Function SendMessageAsAny Lib "user32" Alias "SendMessageA" _
                                            (ByVal hWnd As Long, _
                                            ByVal wMsg As Long, _
                                            ByVal wParam As Long, _
                                            ByRef lParam As Any) As Long
Private Declare Function SendMessageAsString Lib "user32" Alias "SendMessageA" _
                                            (ByVal hWnd As Long, _
                                            ByVal wMsg As Long, _
                                            ByVal wParam As Long, _
                                            ByVal lParam As String) As Long
                                            
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long 'C BOOL


Private Function GetCapSureDevice() As String
    '---------------------------------------------------------------------
    '功能：获取视频设备清单
    '参数：
    '返回：设备清单用";"分开
    '上级函数或过程：
    '下级函数或过程：capGetDriverDescription
    '引用的外部参数：
    '编制人：曾超
    '修改人：
    '---------------------------------------------------------------------
    '获取驱动列表
    Const MAXVIDDRIVERS As Long = 9
    Const CAP_STRING_MAX As Long = 128
    
    Dim Index As Long
    Dim Device As String
    Dim Version As String
    Dim strTmp As String
    
    Device = String$(CAP_STRING_MAX, 0)
    Version = String$(CAP_STRING_MAX, 0)
    For Index = 0 To 8
        If 0 <> capGetDriverDescription(Index, Device, CAP_STRING_MAX, Version, CAP_STRING_MAX) Then
             strTmp = Left(Device, InStr(Device, vbNullChar) - 1) & Left$(Version, InStr(Version, vbNullChar) - 1)
             If Len(Trim(GetCapSureDevice)) > 0 Then
                GetCapSureDevice = GetCapSureDevice & ";"
             End If
             GetCapSureDevice = GetCapSureDevice & strTmp
        End If
    Next
End Function


Private Function ConnCapDevice(ParentWindowWnd As Long, CapDeviceIndex As Integer) As Boolean
    '-----------------------------------------------------------------------------------------
    '功能：连接到设备
    '参数：ParentWindowWnd 父窗体句柄 ; CapDeviceIndex 设备索引号
    '返回：True = 成功 False = 失败
    '上级函数或过程：
    '下级函数或过程：capCreateCaptureWindow;SendMessageAsLong;SendMessageAsAny;SetWindowPos
    '引用的外部参数：hCapWnd
    '编制人：曾超
    '修改人：
    '-----------------------------------------------------------------------------------------
    Dim retVal As Boolean
    Dim capStat As CAPSTATUS
    Dim strTmp() As String
    Dim i  As Integer
    
    hCapWnd = capCreateCaptureWindow("ZLSOFT_CAPTURE", WS_CHILD Or WS_VISIBLE, 0, 0, 5, 5, ParentWindowWnd, 0)
    
    If hCapWnd = 0 Then
        MsgBox "创建采集窗体失败！", vbInformation, gstrSysName
        Exit Function
    End If

    retVal = SendMessageAsLong(hCapWnd, WM_CAP_DRIVER_CONNECT, CapDeviceIndex, 0&)
    
    If retVal = False Then
        MsgBox "连接设备失败！", vbInformation, gstrSysName
        DestroyWindow hCapWnd
        Exit Function
    End If
    
    Call SendMessageAsLong(hCapWnd, WM_CAP_SET_PREVIEWRATE, 66, 0&)
    Call SendMessageAsLong(hCapWnd, WM_CAP_SET_PREVIEW, -(True), 0&)
    
    retVal = SendMessageAsAny(hCapWnd, WM_CAP_GET_STATUS, Len(capStat), capStat)

    Call SetWindowPos(hCapWnd, _
                0&, _
                0&, _
                0&, _
                capStat.uiImageWidth, _
                capStat.uiImageHeight, _
                SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSENDCHANGING)
                
    ConnCapDevice = True

End Function

Private Function SelectCapDevice(CapDeviceIndex As Integer) As Boolean
    '---------------------------------------------------------------------
    '功能：连接到指定设备
    '参数：CapDeviceIndex 设备索引(0--8)
    '返回：True = 成功 False = 失败
    '上级函数或过程：
    '下级函数或过程：SendMessageAsLong;SetWindowPos
    '引用的外部参数：hCapWnd
    '编制人：曾超
    '修改人：
    '---------------------------------------------------------------------
    SelectCapDevice = SendMessageAsLong(hCapWnd, WM_CAP_DRIVER_CONNECT, CapDeviceIndex, 0&)
    
End Function

Private Function ParentWindowResize(ParentWindowWidth As Long, ParentWindowHeight As Long) As Boolean
    '---------------------------------------------------------------------
    '功能：设置显示窗口的位置在父窗体中心
    '参数：ParentWindowWidth 父窗体宽度 ParentWindowHeight 父窗体高度
    '返回：True = 成功 False = 失败
    '上级函数或过程：
    '下级函数或过程：SendMessageAsAny
    '引用的外部参数：hCapWnd
    '编制人：曾超
    '修改人：
    '---------------------------------------------------------------------
    Dim retVal As Boolean
    Dim capStat As CAPSTATUS
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    retVal = SendMessageAsAny(hCapWnd, WM_CAP_GET_STATUS, Len(capStat), capStat)
    
    If retVal Then
        If ParentWindowWidth - capStat.uiImageWidth <= 0 Then
            lngWidth = ParentWindowWidth
        Else
            lngWidth = (ParentWindowWidth - capStat.uiImageWidth) / 2
        End If
        If ParentWindowHeight - capStat.uiImageHeight <= 0 Then
            lngHeight = ParentWindowHeight
        Else
            lngHeight = (ParentWindowHeight - capStat.uiImageHeight) / 2
        End If
        Call SetWindowPos(hCapWnd, 0&, lngWidth, lngHeight, 0&, 0&, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOSENDCHANGING)
    End If
    
    ParentWindowResize = True
    
End Function

Private Function SaveImageFile(SavePath As String) As Boolean
    '---------------------------------------------------------------------
    '功能：保存当前显示的图像
    '参数：SavePath=保存路径
    '返回：True = 成功 False = 失败
    '上级函数或过程：
    '下级函数或过程：SendMessageAsString
    '引用的外部参数：hCapWnd
    '编制人：曾超
    '修改人：
    '---------------------------------------------------------------------
    SaveImageFile = SendMessageAsString(hCapWnd, WM_CAP_FILE_SAVEDIB, 0&, SavePath)
    
End Function

Private Function ViewerFormat() As Boolean
    '---------------------------------------------------------------------
    '功能：显示图像格式
    '参数：
    '返回：True = 成功 False = 失败
    '上级函数或过程：
    '下级函数或过程：SendMessageAsLong
    '引用的外部参数：hCapWnd
    '编制人：曾超
    '修改人：
    '---------------------------------------------------------------------
    ViewerFormat = SendMessageAsLong(hCapWnd, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
    
End Function

Private Function ViewerSource() As Boolean
    '---------------------------------------------------------------------
    '功能：显示图像来源
    '参数：
    '返回：True = 成功 False = 失败
    '上级函数或过程：
    '下级函数或过程：SendMessageAsLong
    '引用的外部参数：hCapWnd
    '编制人：曾超
    '修改人：
    '---------------------------------------------------------------------
    ViewerSource = SendMessageAsLong(hCapWnd, WM_CAP_DLG_VIDEOSOURCE, 0&, 0&)
    
End Function

Public Function ShowEdit(ByVal frmMain As Object, ByRef lngKey As Long, Optional lng病人id As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:  frmMain         调用窗体对象
    '       lngKey          预约登记id
    '返回:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
    mlng病人id = lng病人id
    mlngKey = lngKey
    Set mfrmMain = frmMain
        
    If InitData = False Then Exit Function
    If ReadData(mlngKey, lng病人id) = False Then Exit Function

    Me.Show 1, frmMain
    
    ShowEdit = mblnOK

End Function


Private Function ReadData(ByVal lngKey As Long, ByVal lng病人id As Long) As Boolean
     '------------------------------------------------------------------------------------------------------------------
    '功能:  读取数据
    '参数:  lngKey      体检类型序号
    '返回:  True        读取成功
    '       False       读取失败
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand
    
    gstrSQL = "SELECT Decode(C.病人id,Null,0,0,0,1) AS 照片,A.病人id AS ID,A.姓名,B.门诊号,B.性别,TO_CHAR(B.出生日期,'yyyy-mm-dd') AS 出生日期 " & _
                "FROM 体检人员档案 A,病人信息 B,病人照片 C " & _
                "WHERE C.病人id(+)=A.病人id AND A.体检状态 IN (4,5) AND A.病人id=B.病人id and A.登记id=[1]"
    If lng病人id > 0 Then gstrSQL = gstrSQL & " AND B.病人id=[2]"
    
    gstrSQL = gstrSQL & " Order By B.门诊号"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, lng病人id)
    If rs.BOF = False Then
        Call FillGrid(vsf, rs)
        Call AppendRows(vsf, lnX, lnY)
    End If
        
    
    
    ReadData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  初始化设置
    '返回:  True        初始化成功
    '       False       初始化失败
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    Dim intLoop  As Integer
    Dim strTmp() As String
    
    On Error GoTo errHand
    
    strVsf = "照片,450,1,1,1,;姓名,1080,1,1,1,;门诊号,810,7,1,1,;性别,600,1,1,1,;出生日期,990,1,1,1,"
    
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(0) = flexDTBoolean
        
    Call AppendRows(vsf, lnX, lnY)
    
    '初始化设备
    
    strTmp = Split(GetCapSureDevice(), ";")
    For intLoop = 0 To UBound(strTmp)
        cboDev.AddItem strTmp(intLoop)
    Next
    
    If cboDev.ListCount > 0 Then
        
        cboDev.ListIndex = 0
        Call ConnCapDevice(picFilm.hWnd, cboDev.ListIndex)
        
    End If
    
    InitData = True
    
    Exit Function

errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function SavePhoto(ByVal lng病人id As Long, ByVal strFile As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  保存病人的照片数据
    '参数:  lng病人id
    '       strFile        病人照片文件
    '返回:  保存成功返回True;否则返回False
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    blnTran = True
    gcnOracle.BeginTrans
    gcnOracle.Execute "Delete From 病人照片 Where 病人id=" & lng病人id
    
    gstrSQL = "Select 病人id,照片 From 病人照片 where 病人id=" & lng病人id
    
    rs.Open gstrSQL, gcnOracle, adOpenStatic, adLockOptimistic
    
    If rs.BOF Then
        
        If rs.EOF Then rs.AddNew
        
        rs("病人id").Value = lng病人id
        rs("照片").Value = Null
        rs.Update
        
        If zlDatabase.SavePicture(strFile, rs, "照片") = False Then
        
            ShowSimpleMsg "保存照片有误,请确认文件是否被删除!"
            
            gcnOracle.RollbackTrans
            blnTran = False
            
            Exit Function
            
        End If
        
        rs.Close
    End If
    
    gcnOracle.CommitTrans
    blnTran = False
    
    SavePhoto = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran Then gcnOracle.RollbackTrans
    
End Function

Private Sub cboDev_Click()
    If cboDev.ListIndex <> -1 Then Call SelectCapDevice(cboDev.ListIndex)
End Sub

Private Sub cmdClear_Click()
    Dim blnTran As Boolean
    
    On Error GoTo errHand
    picPhoto.Tag = ""
    picPhoto.Cls
    
    blnTran = True
    gcnOracle.BeginTrans
    gcnOracle.Execute "Delete From 病人照片 Where 病人id=" & Val(vsf.RowData(vsf.Row))
    gcnOracle.CommitTrans
    blnTran = False
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
End Sub

Private Sub cmdFile_Click()
    
    Dim objFile As New FileSystemObject
    Dim strTmp As String
    
    If Val(vsf.RowData(vsf.Row)) = 0 Then Exit Sub
    
    dlg.DialogTitle = "请选择要添加的照片文件"
    dlg.Filter = "图片(*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif"
    
    On Error Resume Next
    
    dlg.Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
    dlg.FileName = ""
    dlg.MaxFileSize = 32767
    dlg.CancelError = True
    dlg.ShowOpen
    
    If Err.Number = 0 And dlg.FileName <> "" Then
        
        picPhoto.Tag = dlg.FileName
        
        On Error GoTo errHand
        
        Call DrawPicture(picPhoto, VB.LoadPicture(picPhoto.Tag), picPhoto.Width, picPhoto.Height)
        Call SavePhoto(Val(vsf.RowData(vsf.Row)), picPhoto.Tag)
        
        vsf.TextMatrix(vsf.Row, 0) = 1
    Else
        Err.Clear
    End If
    
    Exit Sub
    
errHand:
    ShowSimpleMsg "不能打开文件(" & picPhoto.Tag & "),该文件可能正在使用或文件不存在!"
End Sub

Private Sub cmdLoad_Click()
    Dim strTmpFile As String
    
    On Error GoTo errHand
    
    If Val(vsf.RowData(vsf.Row)) = 0 Then Exit Sub
    
    strTmpFile = CreateTmpFile("bmp")
    
    Call SaveImageFile(strTmpFile)
    
    picPhoto.Tag = strTmpFile
    
    Call DrawPicture(picPhoto, VB.LoadPicture(picPhoto.Tag), picPhoto.Width, picPhoto.Height)
    Call SavePhoto(Val(vsf.RowData(vsf.Row)), picPhoto.Tag)
    
    vsf.TextMatrix(vsf.Row, 0) = 1
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub

Private Sub cmdSource_Click()
    Call ViewerSource
End Sub

Private Sub cmdStyle_Click()
    Call ViewerFormat
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
    
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim rs As New ADODB.Recordset
    Dim strTmpFile As String
    Dim objStd As IPictureDisp
    '病人照片
    If NewRow = OldRow Then Exit Sub
    
    picPhoto.Cls
    
    gstrSQL = "Select B.* From 病人照片 B Where B.病人id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(NewRow)))
    
    If rs.BOF = False Then
        strTmpFile = ""
        strTmpFile = ReadPicture(rs, "照片", strTmpFile)
        
        If strTmpFile <> "" Then
            Set objStd = VB.LoadPicture(strTmpFile)
            Call DrawPicture(picPhoto, objStd, objStd.Width, objStd.Height)
        End If
    End If
End Sub

