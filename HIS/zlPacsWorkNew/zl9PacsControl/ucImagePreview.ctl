VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucImagePreview 
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7605
   ScaleHeight     =   3795
   ScaleWidth      =   7605
   ToolboxBitmap   =   "ucImagePreview.ctx":0000
   Begin VB.Timer tmrLoad 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6120
      Top             =   3360
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   6600
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucImagePreview.ctx":0312
            Key             =   "avi"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucImagePreview.ctx":15FC64
            Key             =   "aviDownLoad"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucImagePreview.ctx":2BF5B6
            Key             =   "wav"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucImagePreview.ctx":41EF08
            Key             =   "wavDownLoad"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucImagePreview.ctx":57E85A
            Key             =   "fileDisconet"
         EndProperty
      EndProperty
   End
   Begin zl9PacsControl.ucSplitPage ucPage 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   6210
      _ExtentX        =   10504
      _ExtentY        =   582
      PageCount       =   0
      PageRecord      =   9
      AutoRedrawStyle =   0   'False
   End
   Begin DicomObjects.DicomViewer dcmMiniImage 
      Height          =   3135
      Left            =   105
      TabIndex        =   1
      Top             =   120
      Width           =   7335
      _Version        =   262147
      _ExtentX        =   12938
      _ExtentY        =   5530
      _StockProps     =   35
      BackColor       =   4210752
      UseScrollBars   =   0   'False
      UseMouseWheel   =   -1  'True
   End
   Begin VB.Menu menuPopup 
      Caption         =   "右键菜单"
      Begin VB.Menu mnuSplitPageTool 
         Caption         =   "分页设置(&P)"
      End
      Begin VB.Menu mnuReUpLoad 
         Caption         =   "重新上传(&S)"
      End
   End
End
Attribute VB_Name = "ucImagePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Const M_STR_SELECT_TAG As String = "SELECT"
Private Const M_STR_BORDER_TAG As String = "BORDER"
Private Const M_STR_FAILD_TAG As String = "FAILD"
Private Const CON_INT_DICOMSELECTWIDTH As Integer = 18  'dicom图像黄色选中框左上角黄色小框宽度和高度

Public Enum tQueryLevel
    slAdvice = 0    '医嘱ID
    slStudy = 1     '检查
    slSeries = 2    '序列
    slImage = 3     '图像
    slLocal = 4     '缓存
End Enum

Public Enum TMoveType
    mtLast = 0
    mtNext = 1
    mtFirst = 2
    mtEnd = 3
End Enum

Private mblnIsDock As Boolean  '是否独立窗口，用于分页控件显示
Private mobjFile As New FileSystemObject
Private mstrQueryValue As String         '检查医嘱ID
Private mblnMoved As Boolean             '数据是否被转存
Private mslQueryLevel As tQueryLevel      '图像显示级别
Private mblnQueryTmpRecord As Boolean

Private mblnOnlyLoadReportImage As Boolean     '为True时加载 报告图像 字段中的报告图,反之加载所有报告图
Private mblnIsShowCheckbox As Boolean   '是否显示勾选框
Private mblnEnable As Boolean           '是否可进行编辑
Private mlngBigImageWay As Long         '大图显示方式，0-不显示大图，1-鼠标移动时显示大图，2-鼠标单击时显示大图
Private mlngPreViewTime As Long         '移动预览延时时间
Private mblnIsShowPopup As Boolean      '是否显示右键菜单
Private mtyFileLoadType As FileLoadType
Private mblnIsAutoHidePageControl As Boolean
Private mblnShowPageControl As Boolean
Private mlngFailedLoadCount As Long     '失败加载次数
Private mblnIsLoadReportImage As Boolean '是根据报告图象字段加载的报告图
Private mrsRecord As ADODB.Recordset
Private mlngMouseMoveZoom As Double     '鼠标在图像上移动时，显示大图的放大倍数，如果为0则不显示大图
Private mblnBigImageCtl As Boolean      '大图显示控制，True--按设置的分辨率进行大小控制

Private mblnDo As Boolean      '临时变量，用于暂时屏蔽采集界面图像处理保存报告图功能

Private WithEvents mobjImageProcess As clsImageProcess
Attribute mobjImageProcess.VB_VarHelpID = -1

Private mMultiCols As Long
Private mMultiRows As Long

Private mlngSelectIndex As Long
Private mobjFailedImgs As New Scripting.Dictionary    '下载失败的图像集合

Private mblnClickCheckState As Boolean
Private mintImage As Integer


Public Event OnSelChange(ByVal lngSelectedIndex As Long)
Public Event OnCheckChange(ByVal lngSelectedIndex As Long, ByVal blnSelected As Boolean)
Public Event OnClick(ByVal lngSelectedIndex As Long)
Public Event OnDbClick(ByVal lngSelectedIndex As Long, ByRef blnContinue As Boolean)
Public Event OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)
Public Event OnSaveImage(ByVal dcmImage As DicomImage, ByVal lngImageType As Long)
Public Event OnReUpload()
Public Event AfterSaveStudy()

Private Sub DoOnSelChange(ByVal lngSelectedIndex As Long)
On Error Resume Next
    RaiseEvent OnSelChange(lngSelectedIndex)

    err.Clear
End Sub


Private Sub DoOnCheckChange(ByVal lngSelectedIndex As Long, ByVal blnSelected As Boolean)
On Error Resume Next
    RaiseEvent OnCheckChange(lngSelectedIndex, blnSelected)
    
    err.Clear
End Sub

Private Sub DoOnDbClick(ByVal lngSelectedIndex As Long, ByRef blnContinue As Boolean)
On Error Resume Next

    RaiseEvent OnDbClick(lngSelectedIndex, blnContinue)

    err.Clear
End Sub

Private Sub DoOnClick(ByVal lngSelectedIndex As Long)
On Error Resume Next
    
    RaiseEvent OnClick(lngSelectedIndex)
    
    err.Clear
End Sub

'临时属性，用于暂时屏蔽采集界面图像处理保存报告图功能
Property Get DoShield() As Boolean
    DoShield = mblnDo
End Property

Property Let DoShield(value As Boolean)
    mblnDo = value
End Property

'
Property Get AutoRedrawStyle() As Boolean
    AutoRedrawStyle = AutoRedraw
End Property

Property Let AutoRedrawStyle(value As Boolean)
    AutoRedraw = value
    
    ucPage.AutoRedrawStyle = value
End Property

'鼠标移动到图像上的放大倍数，如果为0则不进行放大
Property Get MouseMoveZoom() As Double
    MouseMoveZoom = mlngMouseMoveZoom
End Property

Property Let MouseMoveZoom(value As Double)
    mlngMouseMoveZoom = value
End Property

'大图显示按按设置的分辨率进行大小控制
Property Get BigImageCtl() As Boolean
    BigImageCtl = mblnBigImageCtl
End Property

Property Let BigImageCtl(value As Boolean)
    mblnBigImageCtl = value
End Property

Property Get OnlyLoadReportImage() As Boolean
    OnlyLoadReportImage = mblnOnlyLoadReportImage
End Property

Property Let OnlyLoadReportImage(value As Boolean)
    mblnOnlyLoadReportImage = value
End Property

'是否显示图像勾选框
Property Get ShowCheckBox() As Boolean
    ShowCheckBox = mblnIsShowCheckbox
End Property

Property Let ShowCheckBox(value As Boolean)
    mblnIsShowCheckbox = value
End Property


'大图显示方式
Property Get BigImageWay() As Long
    BigImageWay = mlngBigImageWay
End Property

Property Let BigImageWay(value As Long)
    mlngBigImageWay = value
End Property

'预览延时时间
Property Get PreViewTime() As Long
    PreViewTime = mlngPreViewTime
End Property

Property Let PreViewTime(value As Long)
    mlngPreViewTime = value
End Property

'是否显示右键菜单
Property Get ShowPopup() As Boolean
    ShowPopup = mblnIsShowPopup
End Property

Property Let ShowPopup(value As Boolean)
    mblnIsShowPopup = value
End Property

'是否可进行编辑
Property Get Enable() As Boolean
    Enable = mblnEnable
End Property

Property Let Enable(value As Boolean)
    mblnEnable = value
End Property

'获取图像呈现组件
Property Get ImgViewer() As Object
    Set ImgViewer = dcmMiniImage
End Property

'图像总数
Property Get ImageTotal() As Long
    ImageTotal = ucPage.RecordCount
End Property

'获取控件句柄
Property Get Handle() As Long
    Handle = UserControl.hWnd
End Property

'文件加载方式
Property Get ImgLoadType() As FileLoadType
    ImgLoadType = mtyFileLoadType
End Property

Property Let ImgLoadType(value As FileLoadType)
    mtyFileLoadType = value
    mnuReUpLoad.Visible = value = Service
End Property

'自动隐藏分页组件，当每页的显示数量小于指定的每页显示数量时
Property Get AutoHidePageControl() As Boolean
    AutoHidePageControl = mblnIsAutoHidePageControl
End Property


Property Let AutoHidePageControl(value As Boolean)
    mblnIsAutoHidePageControl = value
End Property


'查询条件值
Property Get QueryValue() As String
    QueryValue = mstrQueryValue
End Property

Property Let QueryValue(value As String)
    mstrQueryValue = value
End Property


'项目是否被选择
Property Get ImgChecked(Index As Long) As Boolean
    Dim i As Long
    Dim objLabs As DicomLabels
    
    Set objLabs = dcmMiniImage.Images(Index).Labels
    
    ImgChecked = False
    For i = 1 To objLabs.Count
        If objLabs(i).Tag = M_STR_SELECT_TAG Then
            ImgChecked = Not objLabs(i).Transparent And objLabs(i).Visible
            Exit Property
        End If
    Next i
End Property

Property Let ImgChecked(Index As Long, value As Boolean)
    Dim i As Long
    Dim objLabs As DicomLabels
    
    Set objLabs = dcmMiniImage.Images(Index).Labels
    
    For i = 1 To objLabs.Count
        If objLabs(i).Tag = M_STR_SELECT_TAG Then
            objLabs(i).Transparent = Not value
            Call dcmMiniImage.Images(Index).Refresh(False)
            
            Exit Property
        End If
    Next i
End Property

Property Get CellSpacing() As Long
    CellSpacing = dcmMiniImage.CellSpacing
End Property

Property Let CellSpacing(value As Long)
    dcmMiniImage.CellSpacing = value
End Property


'每页图像显示数量
Property Get PageImgCount() As Long
    PageImgCount = ucPage.PageRecord
End Property

Property Let PageImgCount(value As Long)
    ucPage.PageRecord = value
End Property

'当前页数
Property Get PageNumber() As Long
    PageNumber = ucPage.PageNumber
End Property

'背景颜色
Property Get BackColor() As OLE_COLOR
    BackColor = dcmMiniImage.BackColour
End Property


Property Let BackColor(value As OLE_COLOR)
    dcmMiniImage.BackColour = value
End Property

'获取当前选中的索引
Property Get SelectIndex() As Long
    SelectIndex = mlngSelectIndex
End Property


'获取当前选中的图像
Property Get SelectImage() As DicomImage
    Set SelectImage = Nothing
    
    If mlngSelectIndex > 0 And mlngSelectIndex <= dcmMiniImage.Images.Count Then
        Set SelectImage = dcmMiniImage.Images(mlngSelectIndex)
    End If
End Property


'获取当前显示得图像数量
Property Get CurImageCount() As Long
    CurImageCount = dcmMiniImage.Images.Count
End Property


Public Sub RedrawSelf()
'重绘界面
    Call dcmMiniImage.Refresh
    Call ucPage.RedrawSelf
End Sub

Public Sub RefreshFace(ByVal IsDock As Boolean)
'刷新界面控件位置
    mblnIsDock = IsDock
    Call UserControl_Resize
End Sub


Public Sub ShowPageConfig()
'显示分页配置
    Call ShowPageControl
End Sub

Public Sub MovePage(ByVal lngMoveType As TMoveType)
'移动/跳转图像页面
    Select Case lngMoveType
        Case mtLast
            Call ucPage.LastPage
        Case mtNext
            Call ucPage.NextPage
        Case mtFirst
            Call ucPage.FirstPage
        Case mtEnd
            Call ucPage.EndPage
    End Select
End Sub

Public Sub RefreshImage(ByVal slQueryLevel As tQueryLevel, _
    ByVal strQueryValue As String, _
    ByVal blnMoved As Boolean, _
    Optional ByVal blnFoceRefresh As Boolean = False, _
    Optional ByVal blnTmpRecord As Boolean = False)
    
'刷新图像显示
    Dim rsData As ADODB.Recordset
    Dim blnLoadResult As Boolean
    Dim i As Long
    
BUGEX "RefreshImage 1"
    mnuReUpLoad.Enabled = False
    If mstrQueryValue = strQueryValue And Not blnFoceRefresh And slQueryLevel <> slLocal Then
        mslQueryLevel = slQueryLevel
        Exit Sub
    End If
    
    mstrQueryValue = strQueryValue
    mslQueryLevel = slQueryLevel
    mblnQueryTmpRecord = blnTmpRecord
    mblnMoved = blnMoved
    
    If slQueryLevel = slLocal Then
        If mobjFile.FolderExists(strQueryValue) = False Then
            MsgBox "本地缓存目录不存在！", vbExclamation, CON_STR_HINT_TITLE
            Exit Sub
        End If
    End If
    
    ucPage.RecordCount = 0
    mlngSelectIndex = 0
    
BUGEX "RefreshImage 2"
    Call RefreshPageControl
    
BUGEX "RefreshImage 3"
        
    If strQueryValue = "" Then
        '清除图像
        Call ClearCurrentPageImage
        Exit Sub
    End If
    
BUGEX "RefreshImage 4"
    '配置分页组件
    If slQueryLevel = slLocal Then
        Call ConfigPageControlWithLocal(strQueryValue)
    Else
        If mblnOnlyLoadReportImage Then
            Call ConfigRptPageControl(slQueryLevel, strQueryValue, blnTmpRecord)
        Else
            Call ConfigPageControl(slQueryLevel, strQueryValue, blnTmpRecord)
        End If
    End If
    
BUGEX "RefreshImage 5"
    
    '加载图像信息
'    For i = 1 To dcmMiniImage.Images.Count
'        dcmMiniImage.Images(i).BorderColour = vbWhite
'        ImgChecked(i) = False
'    Next

    ChangeImgSelected dcmMiniImage, 1, True
    blnLoadResult = LoadImage(1, ucPage.PageRecord)
    mnuReUpLoad.Enabled = dcmMiniImage.Images.Count > 0
    
BUGEX "RefreshImage 6"
     '绘制图像的各种标注
    Call DrawImageLabels(dcmMiniImage)
        
    Call UserControl_Resize
    
    If blnLoadResult = True Then Call dcmMiniImage.Refresh
    
BUGEX "RefreshImage End"
End Sub

Public Sub RefreshLabelTag()
    '绘制图像的各种标注
    Call DrawImageLabels(dcmMiniImage)
End Sub

Private Sub RefreshPageControl()
'刷新分页组件显示
On Error Resume Next
    If Not mblnIsAutoHidePageControl Then Exit Sub
    
    mblnShowPageControl = IIf(ucPage.RecordCount <= ucPage.PageRecord, False, True)
    ucPage.Visible = mblnShowPageControl
    
    Call UserControl_Resize
    
    err.Clear
End Sub

Public Sub ClearCurrentPageImage()
'清除图像
On Error GoTo errHandle
    mlngSelectIndex = 0
    
    dcmMiniImage.Images.Clear
    dcmMiniImage.MultiColumns = 1
    dcmMiniImage.MultiRows = 1
    
Exit Sub
errHandle:
    MsgboxEx hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE
    err.Clear
End Sub


Private Sub ConfigImgDisplayFormat(ByVal lngPageRecord As Long)
'配置图像显示格式
    Dim iRows As Integer
    Dim iCols As Integer
    
    ResizeRegion lngPageRecord, dcmMiniImage.Width, dcmMiniImage.Height, iRows, iCols
    
    mMultiCols = iCols
    mMultiRows = iRows

    dcmMiniImage.MultiColumns = iCols
    dcmMiniImage.MultiRows = iRows
End Sub

'是否加载失败的图像
Public Function IsFailedImg(Index As Long) As Boolean
    Dim i As Long
    Dim objLabs As DicomLabels
    
    Set objLabs = dcmMiniImage.Images(Index).Labels
    
    IsFailedImg = False
    For i = 1 To objLabs.Count
        If objLabs(i).Tag = M_STR_FAILD_TAG Then
            IsFailedImg = True
            Exit Function
        End If
    Next i
End Function

Private Function SyncDelImage(rsCurImageData As ADODB.Recordset, dcmViewer As DicomViewer) As String
'同步删除的图像
    Dim i As Long
    Dim strImageInstanceUid As String
    Dim blnIsDel As Boolean
    
    SyncDelImage = ""
    blnIsDel = False
    For i = dcmViewer.Images.Count To 1 Step -1
        strImageInstanceUid = dcmViewer.Images(i).InstanceUID
        
        rsCurImageData.Filter = "图像UID ='" & strImageInstanceUid & "'"
        
        If rsCurImageData.RecordCount <= 0 Then
            dcmViewer.Images.Remove (i)
            blnIsDel = True
        Else
            SyncDelImage = SyncDelImage & ";" & strImageInstanceUid & ";"
            
            '重新设置图像的报告图标记，因为刷新视频采集窗体时，如果图像已经纯在，不会重新加载图像信息，所以在这里设置
            dcmViewer.Images(i).Tag.ReportImage = NVL(rsCurImageData!报告图)
        End If
    Next i
    
    If blnIsDel = True Then dcmViewer.Refresh
    
    rsCurImageData.Filter = ""
End Function

Private Function SendDataToservice(ByVal strDataTag As String, ByVal intCommandIdentify As Integer, ByVal strDataFrom As String, fileMsg As TransferFileMsg)
    Dim objServiceHelper As New clsServiceHelper
    
    SendDataToservice = objServiceHelper.SendDataToservice(strDataTag, intCommandIdentify, strDataFrom, fileMsg)
    
    Set objServiceHelper = Nothing
End Function

Private Function LoadViewImageToFaceWithService(rsCurImageData As ADODB.Recordset, dcmViewer As DicomViewer) As Boolean
'通过ZLPacsServerCenter服务加载预览图像到界面
'加载预览图像到界面
    Dim strTmpFile As String
    Dim strCachePath As String
    
    Dim curImage As DicomImage
    
    Dim objFile As New Scripting.FileSystemObject
    
    Dim dcmTag As clsImageTagInf
    Dim strImgInstanceUid As String
    Dim strCurInstanceUids As String
    Dim blnIsAddImage As Boolean
    Dim fileMsg As TransferFileMsg
    Dim blnIsSendOk As Boolean
    
    blnIsAddImage = False
    mlngSelectIndex = 0
    mlngFailedLoadCount = 0
    
    LoadViewImageToFaceWithService = False
    
    If rsCurImageData.RecordCount <= 0 Then
        dcmViewer.Images.Clear
        
        dcmViewer.MultiColumns = 1
        dcmViewer.MultiRows = 1
        
        Exit Function
    End If
    
    strCurInstanceUids = SyncDelImage(rsCurImageData, dcmViewer)
        
    '配置图像显示格式
    If rsCurImageData.RecordCount <> dcmViewer.Images.Count Then
        Call ConfigImgDisplayFormat(IIf(rsCurImageData.RecordCount < ucPage.PageRecord, rsCurImageData.RecordCount, ucPage.PageRecord))
    End If
        
    '创建本地图像缓存目录
    MkLocalDir GetResourceDir
    strCachePath = GetCacheDir
    MkLocalDir strCachePath & objFile.GetParentFolderName(NVL(rsCurImageData("URL")))
    
    Do While Not rsCurImageData.EOF
        '循环加载图像到DicomViewer中
        strImgInstanceUid = Trim(NVL(rsCurImageData!图像UID))
        
        If InStr(strCurInstanceUids, strImgInstanceUid) <= 0 And strImgInstanceUid <> "" Then
            blnIsAddImage = True
            
            '设置音视频的显示文件，如果为音视频文件时，该过程将不从服务器中直接下载数据文件
            If NVL(rsCurImageData!动态图, imgTag) = VIDEOTAG Then
                strTmpFile = GetResourceDir & "Avi.bmp"
            ElseIf NVL(rsCurImageData!动态图, imgTag) = AUDIOTAG Then
                strTmpFile = GetResourceDir & "wav.bmp"
            Else
                strTmpFile = strCachePath & NVL(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", "")
            End If
            
            strTmpFile = Trim(strTmpFile)
            
            blnIsSendOk = True
            
            If Dir(strTmpFile) = vbNullString Then
                '本地缓存图像不存在，将文件数据发送至服务，则使用服务后台下载
                With fileMsg
                    fileMsg.strAdviceId = Val(NVL(rsCurImageData("医嘱ID")))
                    fileMsg.strName = NVL(rsCurImageData("姓名"))
                    fileMsg.strSex = NVL(rsCurImageData("性别"))
                    fileMsg.strAge = NVL(rsCurImageData("年龄"))
                    
                    fileMsg.ftpInfo.strDeviceId = NVL(rsCurImageData("设备号1"))
                    fileMsg.ftpInfo.strFtpDir = NVL(rsCurImageData("Root1"))
                    fileMsg.ftpInfo.strFTPIP = NVL(rsCurImageData("Host1"))
                    fileMsg.ftpInfo.strFTPPwd = NVL(rsCurImageData("Pwd1"))
                    fileMsg.ftpInfo.strFTPUser = NVL(rsCurImageData("User1"))
                    fileMsg.ftpInfo.strSDDir = NVL(rsCurImageData("共享目录1"))
                    fileMsg.ftpInfo.strSDPswd = NVL(rsCurImageData("共享目录密码1"))
                    fileMsg.ftpInfo.strSDUser = NVL(rsCurImageData("共享目录用户名1"))
                    
                    fileMsg.bakFtpInfo.strDeviceId = NVL(rsCurImageData("设备号2"))
                    fileMsg.bakFtpInfo.strFtpDir = NVL(rsCurImageData("Root2"))
                    fileMsg.bakFtpInfo.strFTPIP = NVL(rsCurImageData("Host2"))
                    fileMsg.bakFtpInfo.strFTPPwd = NVL(rsCurImageData("Pwd2"))
                    fileMsg.bakFtpInfo.strFTPUser = NVL(rsCurImageData("User2"))
                    fileMsg.bakFtpInfo.strSDDir = NVL(rsCurImageData("共享目录2"))
                    fileMsg.bakFtpInfo.strSDPswd = NVL(rsCurImageData("共享目录密码2"))
                    fileMsg.bakFtpInfo.strSDUser = NVL(rsCurImageData("共享目录用户名2"))
                    
                    fileMsg.strLocalDir = strTmpFile
                    fileMsg.strFileName = NVL(rsCurImageData("图像UID")) & IIf(mblnIsLoadReportImage, ".jpg", "")
                    fileMsg.strSubDir = NVL(rsCurImageData("URL"))
                    fileMsg.strMediaType = NVL(rsCurImageData!动态图, imgTag)
                End With
                
                If Not SendDataToservice("缩略图", LoadCommand.COMMAND_RPTIMG_DOWNLOAD, "图像下载", fileMsg) Then
                    blnIsSendOk = False
                End If
            End If
            
            If NVL(rsCurImageData!动态图, imgTag) <> VIDEOTAG And NVL(rsCurImageData("动态图"), imgTag) <> AUDIOTAG Then
                '设置图像标记
                Set dcmTag = New clsImageTagInf
                dcmTag.Tag = NVL(rsCurImageData!动态图, imgTag)
                
                If Dir(strTmpFile) = vbNullString Then
                    If Dir(strCachePath & "\fileDisconet.bmp") = vbNullString Then
                        Call SavePicture(imgList.ListImages("fileDisconet").Picture, strCachePath & "\fileDisconet.bmp")
                    End If
                    
                    Set curImage = dcmViewer.Images.AddNew
                    Call curImage.FileImport(strCachePath + "fileDisconet.bmp", "DIB/BMP")
                    curImage.InstanceUID = strImgInstanceUid
                    
                    Dim imgLoadInfo As New DicomLabel
                    Dim iCols As Long, iRows As Long
                    
                    iCols = dcmViewer.MultiColumns
                    iRows = dcmViewer.MultiRows
                    
                    If blnIsSendOk Then
                        imgLoadInfo.Text = "[" + NVL(rsCurImageData!设备名1, NVL(rsCurImageData!设备名2)) + "] 文件下载中..."
                    Else
                        imgLoadInfo.Text = "[" + NVL(rsCurImageData!设备名1, NVL(rsCurImageData!设备名2)) + "] 文件下载请求失败."
                    End If
                                        
                    imgLoadInfo.Width = dcmViewer.Width
                    imgLoadInfo.Height = 20
                    
                    imgLoadInfo.Left = 0
                    imgLoadInfo.Top = dcmViewer.Height / Screen.TwipsPerPixelY / iRows - imgLoadInfo.Height * 2

                    imgLoadInfo.AutoSize = True
                    imgLoadInfo.ShowTextBox = False
                    imgLoadInfo.Font.Size = 12
                    imgLoadInfo.Font.Bold = True
                    imgLoadInfo.ForeColour = vbRed
                    imgLoadInfo.Tag = M_STR_FAILD_TAG
                    
                    Call curImage.Labels.Add(imgLoadInfo)
                    
                    '将失败的图像放入集合中
                    If mobjFailedImgs.Exists(strImgInstanceUid) Then Call mobjFailedImgs.Remove(strImgInstanceUid)
                    Call mobjFailedImgs.Add(strImgInstanceUid, strTmpFile)
                Else
                    Set curImage = ReadViewImage(strTmpFile, dcmViewer)
                End If
                                    
                Set curImage.Tag = dcmTag
                
                With curImage
                    .BorderStyle = 6
                    .BorderWidth = 1
                    .BorderColour = vbWhite
                End With
            Else
                Set curImage = New DicomImage
                    
                If Dir(strTmpFile) = vbNullString Then
                    If NVL(rsCurImageData("动态图"), VIDEOTAG) = VIDEOTAG Then
                        Call SavePicture(imgList.ListImages("avi").Picture, strTmpFile)
                    Else
                        Call SavePicture(imgList.ListImages("wav").Picture, strTmpFile)
                    End If
                End If

                Call curImage.FileImport(strTmpFile, "DIB/BMP")

                Set dcmTag = New clsImageTagInf

                dcmTag.Tag = NVL(rsCurImageData!动态图, VIDEOTAG)
                dcmTag.EncoderName = NVL(rsCurImageData("编码名称"), "")
                dcmTag.CaptureTime = NVL(rsCurImageData("采集时间"))
                
                If NVL(rsCurImageData("动态图"), VIDEOTAG) = VIDEOTAG Then
                    dcmTag.VideoFile = strCachePath & NVL(rsCurImageData("URL")) & ".avi"
                Else
                    dcmTag.VideoFile = strCachePath & NVL(rsCurImageData("URL")) & ".wav"
                End If
                
                dcmTag.RecordTimeLen = Val(NVL(rsCurImageData("录制长度"), "0"))
                
                Set curImage.Tag = dcmTag
                
                curImage.InstanceUID = NVL(rsCurImageData("图像UID"))
                curImage.SeriesUID = NVL(rsCurImageData("序列UID"))
                curImage.StudyUID = NVL(rsCurImageData("检查UID"))
                
                Call ShowAVInf(curImage, dcmTag)
                
                With curImage
                    .BorderStyle = 6
                    .BorderWidth = 1
                    .BorderColour = vbWhite
                End With
                
                Call dcmViewer.Images.Add(curImage)
            End If
            
            '取消自动剪影,因为DicomObjects控件本身对处理剪影有BUG，存在（0028，6100）时，会自动对图像进行剪影，
            '导致晋煤的DSA图像不能正常显示
            '虽然设置图像的mask=0 ,可以取消剪影，但是每次图像被添加到新的Dicomimages之后，自动又将mask设置成1了，
            '这样在程序中无法很好的控制，因此直接去掉（0028，6100）这个属性。
            If Not IsNull(curImage.Attributes(&H28, &H6100).value) Then
                curImage.Attributes.Remove &H28, &H6100
            End If
        End If
        
        rsCurImageData.MoveNext
    Loop
    
    Call UpdateSelectIndex(1)
    
    If Dir(strCachePath & "\fileDisconet.bmp") <> vbNullString Then objFile.DeleteFile (strCachePath & "\fileDisconet.bmp")
    If mobjFailedImgs.Count > 0 Then tmrLoad.Enabled = True '启用Timer，开始加载
    
    LoadViewImageToFaceWithService = IIf(Trim(strCurInstanceUids) <> "" And blnIsAddImage = True, True, False)
End Function

Private Function ReadDicomFile(ByVal strFile As String, dcmImgs As DicomImages) As DicomImage
On Error Resume Next
    Dim curImage As DicomImage
    Dim blnUseUrl As Boolean
    Dim strFileTime As String
    
    blnUseUrl = IIf(InStr(strFile, " ") <= 0, True, False)
    
    If blnUseUrl Then
        'readurl不支持空格
        Set curImage = dcmImgs.ReadURL(strFile)
    Else
        Set curImage = dcmImgs.ReadFile(strFile)
    End If
    
    If err.Number = 0 Then
        Set ReadDicomFile = curImage
        Exit Function
    End If
    
    '2098错误一种是文件不是dicom文件，另一种是存在共享访问错误
    If InStr(err.Description, "sharing violation") > 0 Then
        err.Clear
        strFileTime = Format(Now, "YYMMDD") & GetTickCount
        
        Call FileCopy(strFile, strFile & "_copy_vdat_" & strFileTime)
    
        If blnUseUrl Then
            'readurl不支持空格
            Set curImage = dcmImgs.ReadURL(strFile & "_copy_vdat_" & strFileTime)
        Else
            Set curImage = dcmImgs.ReadFile(strFile & "_copy_vdat_" & strFileTime)
        End If
    
        If err.Number = 0 Then
            Call Kill(strFile & "_copy_vdat_" & strFileTime)
            err.Clear
        Else
            Call Kill(strFile & "_copy_vdat_" & strFileTime)
        End If
    Else
        err.Clear
        Set curImage = dcmImgs.AddNew
        Call curImage.FileImport(strFile, "JPG")
        
        If err.Number <> 0 Then
            err.Clear
            'not a JPG file
            Call curImage.FileImport(strFile, "BMP")
        End If
        
        If err.Number <> 0 Then
            err.Clear
            'not a BMP file
            Call curImage.FileImport(strFile, "AVI")
        End If
        
        If err.Number <> 0 Then
            err.Clear
            'not a AVI file
            Call curImage.FileImport(strFile, "MPG")
        End If
    End If
    
    If err.Number = 0 Then
        Set ReadDicomFile = curImage
        Exit Function
    End If
    
    Set ReadDicomFile = Nothing
    
err.Clear
End Function

Private Function ReadViewImage(ByVal strFile As String, Optional ByRef dcmViewer As DicomViewer = Nothing) As DicomImage
On Error GoTo errHandle
    Dim dImgs As DicomImages
        
    '如果包含_copy_vdat_，说明是临时文件
    If InStr(strFile, "_copy_vdat_") > 0 Then
        Set ReadViewImage = Nothing
        Call Kill(strFile)
        
        Exit Function
    End If
    
    If dcmViewer Is Nothing Then
        Set dImgs = New DicomImages
    Else
        Set dImgs = dcmViewer.Images
    End If
    
    Set ReadViewImage = ReadDicomFile(strFile, dImgs)
    
Exit Function
errHandle:
    Set ReadViewImage = Nothing
End Function


Private Function LoadViewImageToFaceWithNormal(rsCurImageData As ADODB.Recordset, dcmViewer As DicomViewer) As Boolean
'加载预览图像到界面
    Dim strTmpFile As String
    Dim strCachePath As String
    
    Dim curImage As DicomImage
    
    Dim objFile As New Scripting.FileSystemObject
    
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    
    Dim dcmTag As clsImageTagInf
    Dim strImgInstanceUid As String
    Dim strCurInstanceUids As String
    Dim blnIsAddImage As Boolean
    Dim objImgInfo As Object
    
BUGEX "LoadViewImageToFaceWithNormal 1"

    blnIsAddImage = False
    mlngSelectIndex = 0
    
    LoadViewImageToFaceWithNormal = False
        
BUGEX "LoadViewImageToFaceWithNormal 2"
    If rsCurImageData.RecordCount <= 0 Then
        dcmViewer.Images.Clear
        
        dcmViewer.MultiColumns = 1
        dcmViewer.MultiRows = 1
        
        Exit Function
    End If
    
    strCurInstanceUids = SyncDelImage(rsCurImageData, dcmViewer)
        
    '配置图像显示格式
    If rsCurImageData.RecordCount <> dcmViewer.Images.Count Then
        Call ConfigImgDisplayFormat(IIf(rsCurImageData.RecordCount < ucPage.PageRecord, rsCurImageData.RecordCount, ucPage.PageRecord))
    End If
        
    '创建本地图像缓存目录
    strCachePath = GetCacheDir
    MkLocalDir strCachePath & objFile.GetParentFolderName(NVL(rsCurImageData("URL")))
    
BUGEX "LoadViewImageToFaceWithNormal 3"
    Do While Not rsCurImageData.EOF
        '循环加载图像到DicomViewer中
        strImgInstanceUid = Trim(NVL(rsCurImageData!图像UID))
        
        If InStr(strCurInstanceUids, strImgInstanceUid) <= 0 And strImgInstanceUid <> "" Then
            
            blnIsAddImage = True
            
            '设置音视频的显示文件，如果为音视频文件时，该过程将不从服务器中直接下载数据文件
            If NVL(rsCurImageData!动态图, imgTag) = VIDEOTAG Then
                strTmpFile = GetResourceDir & "Avi.bmp"
            ElseIf NVL(rsCurImageData!动态图, imgTag) = AUDIOTAG Then
                strTmpFile = GetResourceDir & "wav.bmp"
            Else
                strTmpFile = strCachePath & NVL(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", "")
            End If
            
            strTmpFile = Replace(Trim(strTmpFile), "/", "\")
            
            If Dir(strTmpFile) = vbNullString Then
                '本地缓存图像不存在，则读取FTP图像
                '建立FTP连接
                If NVL(rsCurImageData("设备号1")) <> vbNullString And Inet1.hConnection = 0 Then
                    If Inet1.FuncFtpConnect(NVL(rsCurImageData("Host1")), NVL(rsCurImageData("User1")), NVL(rsCurImageData("Pwd1"))) = 0 Then
                        If NVL(rsCurImageData("设备号2")) <> vbNullString Then
                            If Inet2.FuncFtpConnect(NVL(rsCurImageData("Host2")), NVL(rsCurImageData("User2")), NVL(rsCurImageData("Pwd2"))) = 0 Then
                                MsgboxEx hWnd, "FTP不能正常连接，请检查网络设置。", vbOKOnly, CON_STR_HINT_TITLE
                                Exit Function
                            End If
                        Else
                            MsgboxEx hWnd, "FTP不能正常连接，请检查网络设置。", vbOKOnly, CON_STR_HINT_TITLE
                            Exit Function
                        End If
                    End If
                End If
                
                If Inet1.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsCurImageData("Root1")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", ""), , hWnd) <> 0 Then
                    '从设备号1提取图像失败，则从设备号2提取图像
                    If NVL(rsCurImageData("设备号2")) <> vbNullString Then
                        If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect NVL(rsCurImageData("Host2")), NVL(rsCurImageData("User2")), NVL(rsCurImageData("Pwd2"))
                            Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsCurImageData("Root2")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", ""), , hWnd)
                        End If
                    End If
                End If
    
    BUGEX "LoadViewImageToFaceWithNormal DCM TmpFile:" & strTmpFile
    
            If Dir(strTmpFile) <> vbNullString Then
                If NVL(rsCurImageData!动态图, imgTag) <> VIDEOTAG And NVL(rsCurImageData("动态图"), imgTag) <> AUDIOTAG Then
                    
    BUGEX "LoadViewImageToFaceWithNormal Dcm ReadURL"
                    
                    Set curImage = ReadViewImage(strTmpFile, dcmViewer)
                    
                    '设置图像标记
                    Set dcmTag = New clsImageTagInf
                    dcmTag.Tag = NVL(rsCurImageData!动态图, imgTag)
                    dcmTag.ReportImage = NVL(rsCurImageData!报告图)
                                       
                    Set curImage.Tag = dcmTag
                    
                    With curImage
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                Else
                    Set curImage = New DicomImage
                    
                    If Dir(strTmpFile) = vbNullString Then
                        If NVL(rsCurImageData("动态图"), VIDEOTAG) = VIDEOTAG Then
                            Call SavePicture(imgList.ListImages("avi").Picture, strTmpFile)
                        Else
                            Call SavePicture(imgList.ListImages("wav").Picture, strTmpFile)
                        End If
                    End If
                    
                    Call curImage.FileImport(strTmpFile, "DIB/BMP")
                    Set dcmTag = New clsImageTagInf
                    
BUGEX "LoadViewImageToFaceWithNormal DCM Set Pro."

                    dcmTag.Tag = NVL(rsCurImageData!动态图, VIDEOTAG)
                    dcmTag.EncoderName = NVL(rsCurImageData("编码名称"), "")
                    dcmTag.CaptureTime = NVL(rsCurImageData("采集时间"))
                    dcmTag.ReportImage = NVL(rsCurImageData!报告图)
                    
                    If NVL(rsCurImageData("动态图"), VIDEOTAG) = VIDEOTAG Then
                        dcmTag.VideoFile = strCachePath & NVL(rsCurImageData("URL")) & ".avi"
                    Else
                        dcmTag.VideoFile = strCachePath & NVL(rsCurImageData("URL")) & ".wav"
                    End If
                    
                    dcmTag.RecordTimeLen = Val(NVL(rsCurImageData("录制长度"), "0"))
                    
'                        '如果是视频录像文件，则在播放时进行下载
'                        If Trim(dcmTag.VideoFile) <> "" And Dir(dcmTag.VideoFile) <> "" Then
'                            Name dcmTag.VideoFile As dcmTag.VideoFile & ".avi"
'                        End If
                    
                    Set curImage.Tag = dcmTag
                    
                    curImage.InstanceUID = NVL(rsCurImageData("图像UID"))
                    curImage.SeriesUID = NVL(rsCurImageData("序列UID"))
                    curImage.StudyUID = NVL(rsCurImageData("检查UID"))
                    
                    Call ShowAVInf(curImage, dcmTag)
                    
                    With curImage
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                    
    BUGEX "LoadViewImageToFaceWithNormal DCM AddImage"
                    Call dcmViewer.Images.Add(curImage)
                End If
                
                
                '取消自动剪影,因为DicomObjects控件本身对处理剪影有BUG，存在（0028，6100）时，会自动对图像进行剪影，
                '导致晋煤的DSA图像不能正常显示
                '虽然设置图像的mask=0 ,可以取消剪影，但是每次图像被添加到新的Dicomimages之后，自动又将mask设置成1了，
                '这样在程序中无法很好的控制，因此直接去掉（0028，6100）这个属性。
                If Not IsNull(curImage.Attributes(&H28, &H6100).value) Then
                    curImage.Attributes.Remove &H28, &H6100
                End If
            End If
        End If
        
        rsCurImageData.MoveNext
    Loop
    
    Call UpdateSelectIndex(1)
    
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect
    
    LoadViewImageToFaceWithNormal = IIf(Trim(strCurInstanceUids) <> "" And blnIsAddImage = True, True, False)
    
BUGEX "LoadViewImageToFaceWithNormal End"
End Function

Private Function LoadViewImageToFaceFromLocal(rsCurImageData As ADODB.Recordset, dcmViewer As DicomViewer) As Boolean
    Dim strTmpFile As String
    Dim curImage As DicomImage
    Dim dcmTag As clsImageTagInf
    
On Error GoTo ErrorHand

    If rsCurImageData.RecordCount <= 0 Then
        dcmViewer.Images.Clear
        
        dcmViewer.MultiColumns = 1
        dcmViewer.MultiRows = 1
        
        Exit Function
    End If
    
    '配置图像显示格式
    If rsCurImageData.RecordCount <> dcmViewer.Images.Count Then
        Call ConfigImgDisplayFormat(IIf(rsCurImageData.RecordCount < ucPage.PageRecord, rsCurImageData.RecordCount, ucPage.PageRecord))
    End If
  
    Do While Not rsCurImageData.EOF
        strTmpFile = Trim(NVL(rsCurImageData!路径))
        
        Set curImage = ReadViewImage(strTmpFile, dcmMiniImage)
        
        '设置图像标记
        Set dcmTag = New clsImageTagInf
        dcmTag.Tag = imgTag
        dcmTag.FilePath = strTmpFile
                            
        Set curImage.Tag = dcmTag
        
        With curImage
            .BorderStyle = 6
            .BorderWidth = 1
            .BorderColour = vbWhite
        End With
        
        rsCurImageData.MoveNext
    Loop
    
     Call UpdateSelectIndex(1)
     
     LoadViewImageToFaceFromLocal = True
     Exit Function
ErrorHand:
    LoadViewImageToFaceFromLocal = False
    BUGEX "LoadViewImageToFaceFromLocal err = " & err.Description
End Function


Public Sub PlayMedia(ByVal lngMediaIndex As Long)
'播放指定索引处的媒体

End Sub

Private Sub ConfigPageControlWithLocal(ByVal strQueryPath As String)
    Dim objFile As File
    
    If mobjFile.FolderExists(strQueryPath) = False Then Exit Sub
    
    ucPage.RecordCount = mobjFile.GetFolder(strQueryPath).Files.Count
End Sub

Private Sub ConfigRptPageControl(ByVal slQueryLevel As tQueryLevel, ByVal strSearchValue As String, ByVal blnTmpRecord As Boolean)
'配置分页控件
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    
    strSQL = "Select Count(B.Column_Value) 返回值 From 影像检查记录 A, Table(Cast(f_Str2list(Replace(A.报告图象,';',',')) As zlTools.t_Strlist)) B Where 医嘱ID = [1]"
    '如果查询临时记录，则需要将查询表替换为临时存储数据的表
    If blnTmpRecord Then
        strSQL = Replace(strSQL, "影像检查", "影像临时")
    Else
        If mblnMoved Then
            strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
            strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
        End If
    End If
        
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询图像数量", strSearchValue)
    If rsData.RecordCount > 0 Then lngRecordCount = NVL(rsData!返回值)
    
    If lngRecordCount <= 0 Then
        Select Case slQueryLevel
            Case slAdvice
                strSQL = "select count(1)  as 返回值 from 影像检查图象 a, 影像检查序列 b, 影像检查记录 c where a.序列UID=b.序列UID and b.检查UID=c.检查UID and nvl(a.动态图,0)=0 and c.医嘱ID=[1]"
            Case slStudy
                strSQL = "select count(1)  as 返回值 from 影像检查图象 a, 影像检查序列 b where a.序列UID=b.序列UID and nvl(a.动态图,0)=0 and b.检查UID=[1]"
            Case slSeries
                strSQL = "select count(1)  as 返回值 from 影像检查图象  where nvl(动态图,0)=0 and 序列UID=[1]"
            Case slImage
                strSQL = "select count(1)  as 返回值 from 影像检查图象  where nvl(动态图,0)=0 and 图像UID=[1]"
        End Select
        
        '如果查询临时记录，则需要将查询表替换为临时存储数据的表
        If blnTmpRecord Then
            strSQL = Replace(strSQL, "影像检查", "影像临时")
        Else
            If mblnMoved Then
                strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
                strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
            End If
        End If
    
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询图像数量", strSearchValue)
        
        If rsData.RecordCount > 0 Then
            lngRecordCount = NVL(rsData!返回值)
        Else
            lngRecordCount = 0
        End If
    End If
    
    ucPage.RecordCount = lngRecordCount
    
    Call RefreshPageControl
End Sub

Private Sub ConfigPageControl(ByVal slQueryLevel As tQueryLevel, ByVal strSearchValue As String, ByVal blnTmpRecord As Boolean)
'配置分页控件
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    
    Select Case slQueryLevel
        Case slAdvice
            strSQL = "select count(1)  as 返回值 from 影像检查图象 a, 影像检查序列 b, 影像检查记录 c where a.序列UID=b.序列UID and b.检查UID=c.检查UID and c.医嘱ID=[1]"
        Case slStudy
            strSQL = "select count(1)  as 返回值 from 影像检查图象 a, 影像检查序列 b where a.序列UID=b.序列UID and b.检查UID=[1]"
        Case slSeries
            strSQL = "select count(1)  as 返回值 from 影像检查图象  where  序列UID=[1]"
        Case slImage
            strSQL = "select count(1)  as 返回值 from 影像检查图象  where  图像UID=[1]"
    End Select
    
    '如果查询临时记录，则需要将查询表替换为临时存储数据的表
    If blnTmpRecord Then
        strSQL = Replace(strSQL, "影像检查", "影像临时")
    Else
        If mblnMoved Then
            strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
            strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
        End If
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询图像数量", strSearchValue)
    
    If rsData.RecordCount > 0 Then
        lngRecordCount = NVL(rsData!返回值)
    Else
        lngRecordCount = 0
    End If
    
    ucPage.RecordCount = lngRecordCount
    
    Call RefreshPageControl
End Sub

Private Function GetImageViewDataFromLocal(ByVal lngCurPage As Long, ByVal lngPageRecord As Long) As ADODB.Recordset
    Dim objFile As File
    Dim strDatas() As String
    Dim lngTmpCount As Long
    Dim rsData As New ADODB.Recordset
    Dim lngStartRecord As Long, lngEndRecord As Long
    
    If mobjFile.FolderExists(mstrQueryValue) = False Then Exit Function
    If mobjFile.GetFolder(mstrQueryValue).Files.Count <= 0 Then Exit Function
    
    rsData.Fields.Append "路径", adVarChar, 4000
    rsData.Open
    
    For Each objFile In mobjFile.GetFolder(mstrQueryValue).Files
        lngTmpCount = lngTmpCount + 1
        ReDim Preserve strDatas(lngTmpCount - 1) As String
        strDatas(UBound(strDatas)) = objFile.Path
    Next
    
    lngStartRecord = (lngCurPage - 1) * lngPageRecord + 1
    lngEndRecord = lngCurPage * lngPageRecord
    
    For lngTmpCount = lngStartRecord - 1 To lngEndRecord - 1
        If UBound(strDatas) >= lngTmpCount Then
            rsData.AddNew
            rsData!路径 = strDatas(lngTmpCount)
            rsData.Update
        Else
            Exit For
        End If
    Next
    
    If rsData.RecordCount > 0 Then rsData.MoveFirst
    Set GetImageViewDataFromLocal = rsData
End Function

Private Function GetImageRptData(ByVal lngOrderID As Long, ByVal lngCurPage As Long, ByVal lngPageRecord As Long) As ADODB.Recordset
'根据报告图象 字段获取相关图像
    Dim strSQL As String
    Dim lngStartRecord As Long
    Dim lngEndRecord As Long
    
    strSQL = "Select rownum As 顺序号,a.医嘱id,a.姓名,a.性别,a.年龄, rownum As 图像号,Replace(Trim(D.Column_Value),'.jpg','') as 图像UID, A.检查UID, " & _
            "'' As 序列UID, 0 as 动态图,'' as 编码名称,'' as 采集时间, '' as 录制长度, '' as 报告图," & _
            "B.FTP用户名 As User1,B.FTP密码 As Pwd1,B.IP地址 As Host1,'/'||B.Ftp目录||'/' As Root1, " & _
            "B.共享目录 as 共享目录1,B.共享目录用户名 as 共享目录用户名1,B.共享目录密码 as 共享目录密码1, " & _
            "Decode(A.接收日期,Null,'',to_Char(A.接收日期,'YYYYMMDD')||'/') ||A.检查UID||'/'||Replace(Trim(D.Column_Value),'.jpg','') As URL,B.设备号 as 设备号1, B.设备名 as 设备名1, " & _
            "C.FTP用户名 As User2,C.FTP密码 As Pwd2,C.IP地址 As Host2,'/'||C.Ftp目录||'/' As Root2, " & _
            "C.共享目录 as 共享目录2,C.共享目录用户名 as 共享目录用户名2,C.共享目录密码 as 共享目录密码2,C.设备号 as 设备号2, C.设备名 as 设备名2 " & _
            "From 影像检查记录 A, 影像设备目录 B, 影像设备目录 C, Table(Cast(f_Str2list(A.报告图象,';') As zlTools.t_Strlist)) D " & _
            "Where A.位置一 = B.设备号(+) And A.位置二 = C.设备号(+) And A.医嘱id = [1]"
            
    If mblnMoved = True Then strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    
    lngStartRecord = (lngCurPage - 1) * lngPageRecord + 1
    lngEndRecord = lngCurPage * lngPageRecord
    
    strSQL = "select * from (" & strSQL & " order by 序列UID, 图像号) where 顺序号>=" & lngStartRecord & " and 顺序号<=" & lngEndRecord
    
    Set GetImageRptData = zlDatabase.OpenSQLRecord(strSQL, "提取报告图像", lngOrderID)
End Function

Private Function GetImageViewData(ByVal slQueryLevel As tQueryLevel, ByVal strSearchValue As String, _
    ByVal lngCurPage As Long, ByVal lngPageRecord As Long, ByVal blnTmpRecord As Boolean) As ADODB.Recordset
'获取预览图像数据
'intSearchType:0-按检查uid搜索,1-按序列UID搜索,2-按图像UID搜索

    Dim strSQL As String
    Dim lngStartRecord As Long
    Dim lngEndRecord As Long
    
    strSQL = "Select rownum as 顺序号,[2] 医嘱id,c.姓名,c.性别,c.年龄, A.图像号,D.FTP用户名 As User1,D.FTP密码 As Pwd1," & _
            "D.IP地址 As Host1,'/'||D.Ftp目录||'/' As Root1,D.共享目录 as 共享目录1,D.共享目录用户名 as 共享目录用户名1,D.共享目录密码 as 共享目录密码1," & _
            "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/') " & _
            "||C.检查UID||'/'||A.图像UID As URL,d.设备号 as 设备号1, D.设备名 As 设备名1," & _
            "E.FTP用户名 As User2,E.FTP密码 As Pwd2,A.报告图," & _
            "E.IP地址 As Host2,'/'||E.Ftp目录||'/' As Root2,E.共享目录 as 共享目录2,E.共享目录用户名 as 共享目录用户名2,E.共享目录密码 as 共享目录密码2," & _
            "E.设备号 as 设备号2, E.设备名 As 设备名2, A.图像UID,C.检查UID,B.序列UID,A.动态图,A.编码名称,A.采集时间, A.录制长度 " & _
            "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
            "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+)" & IIf(mblnOnlyLoadReportImage, " And nvl(A.动态图,0) = 0 ", "")
    
    If blnTmpRecord Then
        strSQL = Replace(strSQL, "影像检查", "影像临时")
    Else
        If mblnMoved Then
            strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
            strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
            strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
        End If
    End If

    Select Case slQueryLevel
        Case slAdvice
            strSQL = "select * from (" & strSQL & " and C.医嘱ID=[1])"
        Case slStudy
            strSQL = "select * from (" & strSQL & " and C.检查UID=[1])"
        Case slSeries
            strSQL = "select * from (" & strSQL & " and B.序列UID=[1])"
        Case slImage
            strSQL = "select * from (" & strSQL & " and A.图像UID=[1])"
    End Select
    
    lngStartRecord = (lngCurPage - 1) * lngPageRecord + 1
    lngEndRecord = lngCurPage * lngPageRecord
    
    strSQL = "select * from (" & strSQL & " order by 序列UID, 图像号) where 顺序号>=" & lngStartRecord & " and 顺序号<=" & lngEndRecord
    
    Set GetImageViewData = zlDatabase.OpenSQLRecord(strSQL, "查询图像信息", strSearchValue, IIf(mblnQueryTmpRecord, "-1", mstrQueryValue))
End Function

Public Sub AddImage(Img As Object, Optional objImgTag As Object = Nothing)
'增加图像
    Dim i As Long
    
    If dcmMiniImage.Images.Count < ucPage.PageRecord Then
        Call ConfigImgDisplayFormat(dcmMiniImage.Images.Count + 1)
        
        Call dcmMiniImage.Images.Add(Img)
    Else
        '移动图像
        For i = 2 To dcmMiniImage.Images.Count
            Call dcmMiniImage.Images.Move(i, i - 1)
            dcmMiniImage.Images(i - 1).BorderColour = vbWhite
        Next i
        
        Call dcmMiniImage.Images.Remove(dcmMiniImage.Images.Count)
        dcmMiniImage.Images.Add Img
    End If
    
    '设置选中的边框颜色
    With dcmMiniImage.Images(dcmMiniImage.Images.Count)
        .BorderWidth = 1
        .BorderStyle = 6
        .BorderColour = vbRed
        
        If Not objImgTag Is Nothing Then
            Set .Tag = objImgTag
        End If
        
        If Not .Tag Is Nothing Then
            If UCase(TypeName(.Tag)) = UCase("clsImageTagInf") Then Call ShowAVInf(dcmMiniImage.Images(dcmMiniImage.Images.Count), .Tag)
        End If

    End With
    
    '绘制图像的各种标注
    Call DrawImageLabels(dcmMiniImage)
    
    Call UpdateSelectIndex(dcmMiniImage.Images.Count)
    Call UpdateImageCount(1)
    mnuReUpLoad.Enabled = dcmMiniImage.Images.Count > 0
End Sub


Private Sub ShowAVInf(Img As DicomImage, objImgTag As Object)
'显示音视频信息
    If objImgTag.Tag = VIDEOTAG Or objImgTag.Tag = AUDIOTAG Then
        Call AddVideoLabelToDicomImage(Img, _
        IIf(objImgTag.Tag = VIDEOTAG, "录像时间：", "录音时间：") & objImgTag.CaptureTime, _
        IIf(objImgTag.Tag = VIDEOTAG, "录像长度：", "录音长度：") & objImgTag.RecordTimeLen & " 秒", _
        "编码名称：" & objImgTag.EncoderName)
    End If
End Sub

Public Sub DeleteImage(ByVal lngImgIndex As Long, Optional blMovePage As Boolean = True, Optional blMustMovePage As Boolean = False)
'删除图像 blMovePage:是否判断自动翻页 blMustMovePage是否强制翻页
    Dim i As Long
    Dim lngCurPageCount As Long
        
    
    Call dcmMiniImage.Images.Remove(lngImgIndex)
    
    For i = lngImgIndex + 1 To dcmMiniImage.Images.Count
        Call dcmMiniImage.Move(i, i - 1)
    Next i

    Call dcmMiniImage.Refresh
    
    If lngImgIndex <= dcmMiniImage.Images.Count Then
        Call UpdateSelectIndex(lngImgIndex)
    Else
        Call UpdateSelectIndex(dcmMiniImage.Images.Count)
    End If
    
    
    lngCurPageCount = ucPage.PageCount
    
    Call UpdateImageCount(-1)
        
    If lngCurPageCount > ucPage.PageCount Then
        If blMovePage Then
            Call ucPage.MovePage(ucPage.PageNumber)
            If dcmMiniImage.Images.Count > 0 Then Call UpdateSelectIndex(1)
        End If
    Else
        If blMovePage And blMustMovePage Then

            Call ucPage.MovePage(ucPage.PageNumber)
            If dcmMiniImage.Images.Count > 0 Then Call UpdateSelectIndex(1)
        End If
    End If
    
    For i = 1 To dcmMiniImage.Images.Count
        If i <> mlngSelectIndex Then dcmMiniImage.Images(i).BorderColour = vbWhite
    Next
    
    mnuReUpLoad.Enabled = dcmMiniImage.Images.Count > 0
End Sub

Private Sub UpdateSelectIndex(ByVal lngSelectIndex As Long)
'配置图像的选中索引
    Dim blnIsValidIndex As Boolean
    
    blnIsValidIndex = IIf(lngSelectIndex > 0 And lngSelectIndex <= dcmMiniImage.Images.Count, True, False)
    
    If Not blnIsValidIndex Then Exit Sub

    If blnIsValidIndex Then dcmMiniImage.Images(lngSelectIndex).BorderColour = vbRed
    If mlngSelectIndex = lngSelectIndex Then Exit Sub

    If mlngSelectIndex > 0 And mlngSelectIndex <= dcmMiniImage.Images.Count Then
        dcmMiniImage.Images(mlngSelectIndex).BorderColour = vbWhite
    End If

    mlngSelectIndex = lngSelectIndex
    
    '执行索引改变事件
    Call DoOnSelChange(mlngSelectIndex)
End Sub


Private Sub UpdateImageCount(ByVal lngValue As Long)
    ucPage.RecordCount = ucPage.RecordCount + lngValue
    
    Call RefreshPageControl
End Sub


Public Function SelectedCount() As Long
'获取选择的图像数量
    Dim i As Long
    Dim j As Long
    Dim lngCount As Long
    Dim objLabs As DicomLabels
    
    
    lngCount = 0
    For i = 1 To dcmMiniImage.Images.Count
        Set objLabs = dcmMiniImage.Images(i).Labels
        
        For j = 1 To objLabs.Count
            If objLabs(j).Tag = M_STR_SELECT_TAG Then
                If Not objLabs(j).Transparent Then lngCount = lngCount + 1
                Exit For
            End If
        Next j
    Next i
    
    SelectedCount = lngCount
End Function



Private Sub dcmMiniImage_Click()
On Error GoTo errHandle
    Dim i As Integer
    
    If mlngSelectIndex <= 0 Or mlngSelectIndex > dcmMiniImage.Images.Count Then Exit Sub
    
    Call DoOnClick(mlngSelectIndex)
Exit Sub
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Sub

Private Sub dcmMiniImage_DblClick()
On Error GoTo errHandle
    Dim blnContinue As Boolean
    
    If dcmMiniImage.Images.Count <= 0 Then Exit Sub
    If mlngSelectIndex <= 0 Then Exit Sub

    blnContinue = True
    
    If mlngBigImageWay = 1 Then  '关闭大图显示
        ReleaseCapture      '解锁鼠标
'        frmShowImg.HideMe
    End If
    
    Call DoOnDbClick(mlngSelectIndex, blnContinue)
    
    ImgChecked(mlngSelectIndex) = mblnClickCheckState
    
    If Not blnContinue Then Exit Sub

    
    If dcmMiniImage.MultiColumns = 1 And dcmMiniImage.MultiRows = 1 Then
        dcmMiniImage.MultiColumns = mMultiCols
        dcmMiniImage.MultiRows = mMultiRows
        dcmMiniImage.CurrentIndex = 1
    Else
        mMultiCols = dcmMiniImage.MultiColumns
        mMultiRows = dcmMiniImage.MultiRows
        
        dcmMiniImage.MultiColumns = 1
        dcmMiniImage.MultiRows = 1
        
        dcmMiniImage.CurrentIndex = mlngSelectIndex
    End If
    
    Exit Sub
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Sub


Private Sub dcmMiniImage_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHandle
    Dim lngSelectIndex As Long
    Dim i As Long
    
    If dcmMiniImage.Images.Count <= 0 Then Exit Sub
    
    If KeyCode = 37 Then '<-
        lngSelectIndex = mlngSelectIndex - 1
        If lngSelectIndex <= 0 Then Exit Sub
    ElseIf KeyCode = 38 Then
        lngSelectIndex = mlngSelectIndex - dcmMiniImage.MultiColumns
        If lngSelectIndex <= 0 Then Exit Sub
    ElseIf KeyCode = 39 Then
        lngSelectIndex = mlngSelectIndex + 1
        If lngSelectIndex > dcmMiniImage.Images.Count Then Exit Sub
    ElseIf KeyCode = 40 Then
        lngSelectIndex = mlngSelectIndex + dcmMiniImage.MultiColumns
        If lngSelectIndex > dcmMiniImage.Images.Count Then Exit Sub
    Else
        Exit Sub
    End If
    
    Call UpdateSelectIndex(lngSelectIndex)
    
    If mblnEnable Then
        For i = 1 To dcmMiniImage.Images.Count
            If i <> lngSelectIndex Then
                ImgChecked(i) = False
            Else
                ImgChecked(i) = True
            End If
        Next i
    End If
    
    Call DoOnClick(lngSelectIndex)
Exit Sub
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Sub

Private Sub dcmMiniImage_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error Resume Next
    Dim i As Long
    Dim objLabs As DicomLabels
    Dim lngImgIndex As Long
    Dim blnClickCheck As Boolean
    
    mblnClickCheckState = ImgChecked(mlngSelectIndex)

        lngImgIndex = dcmMiniImage.ImageIndex(X, Y)
        
        Call UpdateSelectIndex(lngImgIndex)
        
        If mlngSelectIndex > 0 And mlngSelectIndex <= dcmMiniImage.Images.Count Then
            
            If mblnEnable Then
                            
                '设置选择框状态
                blnClickCheck = False
                Set objLabs = dcmMiniImage.LabelHits(X, Y, False, True, True)
                For i = 1 To objLabs.Count
                    If objLabs(i).Tag = M_STR_SELECT_TAG And objLabs(i).Visible Then
                        '若objLabs(i).Visible=false，说明选中框已经被隐藏，不做选中处理
                        blnClickCheck = True

                        objLabs(i).Transparent = Not objLabs(i).Transparent
            
                        Call dcmMiniImage.Images(lngImgIndex).Refresh(False)
                        
                        '触发图像勾选事件
                        Call DoOnCheckChange(mlngSelectIndex, Not objLabs(i).Transparent)
                        
                        Exit For
                    End If
                Next i
                
                            '先取消选择
                If Shift = 0 Then
                    If blnClickCheck = False Then
                        If Button = 2 Then
                            If Not ImgChecked(lngImgIndex) Then
                                ChangeImgSelected dcmMiniImage, lngImgIndex, False
                            End If
                        ElseIf Button = 1 Then
                            ChangeImgSelected dcmMiniImage, lngImgIndex, False
                            ImgChecked(lngImgIndex) = Not ImgChecked(lngImgIndex)
                        End If
                    End If
                Else
                    If blnClickCheck = False And Button = 1 Then ImgChecked(lngImgIndex) = Not ImgChecked(lngImgIndex)
                End If
                
            End If
        End If

    
    RaiseEvent OnMouseDown(Button, Shift, X, Y)
End Sub

Private Sub ChangeImgSelected(dcmImage As DicomViewer, lngImage As Long, blnChaBorderColor As Boolean)
    Dim i As Long
       
    For i = 1 To dcmImage.Images.Count
    
        If i <> lngImage Then ImgChecked(i) = False '改变Check框的选中
        
        If blnChaBorderColor Then dcmMiniImage.Images(i).BorderColour = vbWhite '改变边框颜色
    Next
End Sub

Private Sub dcmMiniImage_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)

On Error GoTo errHandle
    Dim blnShowImg As Boolean
    Dim intCurrImg As Integer
    
    If mlngBigImageWay <> 1 Then Exit Sub
    
    '判断是否需要显示图像
    If (0 <= X * Screen.TwipsPerPixelX) And (X * Screen.TwipsPerPixelX <= dcmMiniImage.Width) And _
       (0 <= Y * Screen.TwipsPerPixelY) And (Y * Screen.TwipsPerPixelY <= dcmMiniImage.Height) Then
        blnShowImg = True
    End If
    
    If blnShowImg Then        '显示图像
        SetCapture dcmMiniImage.hWnd    '锁定鼠标
        
        intCurrImg = dcmMiniImage.ImageIndex(X, Y)
        
        
        If intCurrImg <> 0 And intCurrImg <> mintImage Then
            If dcmMiniImage.Images(intCurrImg).Tag.Tag <> VIDEOTAG And dcmMiniImage.Images(intCurrImg).Tag.Tag <> AUDIOTAG Then
            '加载图像并显示
            
                If mobjImageProcess Is Nothing Then
                    Set mobjImageProcess = New clsImageProcess
                End If
                
                mobjImageProcess.ShowImageProcess mstrQueryValue, dcmMiniImage.Images(intCurrImg), ucPage.PageRecord * (ucPage.PageNumber - 1) + intCurrImg, Me, mblnMoved, mslQueryLevel, 1, mlngPreViewTime, mblnDo
    '            frmShowImg.ShowMe GetBigImage(dcmMiniImage.Images(intCurrImg)), Me, 1, 0, 0, BigImageCtl, mlngMouseMoveZoom
                
            End If
            
        End If
        mintImage = intCurrImg
    Else
        ReleaseCapture
        mintImage = 0
    End If
    
    RaiseEvent OnMouseMove(Button, Shift, X, Y)
Exit Sub
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Sub

Private Function GetBigImage(dcmImg As DicomImage) As DicomImage
    
    Set GetBigImage = dcmImg.SubImage(0, 0, dcmImg.SizeX, dcmImg.SizeY, 1, dcmImg.Frame)
     
    GetBigImage.Labels.Clear
    GetBigImage.BorderColour = vbWhite
End Function

Private Sub dcmMiniImage_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errHandle
    Dim curPointer As POINTAPI
    Dim i As Integer
    
    If mlngBigImageWay = 1 Then  '关闭大图显示
        ReleaseCapture      '解锁鼠标
'        frmShowImg.HideMe
    End If
    
    If Button = 2 And mblnIsShowPopup Then
        '显示右键菜单
        Call GetCursorPos(curPointer)
        
        Call ScreenToClient(hWnd, curPointer)  'ScreenToClient方法使用的单位为像素值
        Call PopupMenu(menuPopup, 0, ScaleX(curPointer.X, vbPixels, vbTwips), ScaleY(curPointer.Y, vbPixels, vbTwips))
        
    Else
        '显示大图
        If mlngBigImageWay = 2 And Button = 1 Then
            
            If dcmMiniImage.Images.Count > 0 Then

                i = dcmMiniImage.ImageIndex(X, Y)
                If i = 0 Then i = 1
                
                If dcmMiniImage.Images(i).Tag.Tag <> VIDEOTAG And dcmMiniImage.Images(i).Tag.Tag <> AUDIOTAG Then
                '加载图像并显示
                
                    If mobjImageProcess Is Nothing Then
                        Set mobjImageProcess = New clsImageProcess
                    End If
                    
                    mobjImageProcess.ShowImageProcess mstrQueryValue, dcmMiniImage.Images(i), ucPage.PageRecord * (ucPage.PageNumber - 1) + i, Me, mblnMoved, mslQueryLevel, 1, 0, mblnDo
        '            frmShowImg.ShowMe GetBigImage(dcmMiniImage.Images(intCurrImg)), Me, 1, 0, 0, BigImageCtl, mlngMouseMoveZoom
                End If
'                frmShowImg.ShowMe GetBigImage(dcmMiniImage.Images(i)), Me, 2, 0, 0, BigImageCtl
            End If
        End If
    End If
    
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
errHandle:
End Sub

Private Sub dcmMiniImage_MouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)
On Error GoTo errHandle
    If Delta > 0 Then
        Call ucPage.LastPage
    Else
        Call ucPage.NextPage
    End If
    
    RaiseEvent OnMouseWheel(Shift, Delta, X, Y)
Exit Sub
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Sub

Private Sub mnuReUpLoad_Click()
'重新上传选择的文件
On Error GoTo errHandle
    If mlngSelectIndex > 0 And mlngSelectIndex <= dcmMiniImage.Images.Count Then
        RaiseEvent OnReUpload
    End If
    
    Exit Sub
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Sub

Private Sub mnuSplitPageTool_Click()
'显示分页工具栏
    Call ShowPageControl
End Sub

Private Sub mobjImageProcess_AfterSaveStady()
    RaiseEvent AfterSaveStudy
End Sub

Private Sub mobjImageProcess_OnSaveImage(ByVal dcmImage As DicomObjects.DicomImage, ByVal lngImageType As Long)
    RaiseEvent OnSaveImage(dcmImage, lngImageType)
End Sub

Private Sub mobjImageProcess_OnUnload()
    Set mobjImageProcess = Nothing
End Sub

Private Sub tmrLoad_Timer()
'使用后台服务下载图像时，可能延迟，故在Timer中加载之前未加载的图像
    Dim i As Long, j As Long
    Dim strTmpFile As String
    Dim strTmpKey, dcmTag As Object
    Dim objTmpImg As DicomImage
    Dim iCols As Long, iRows As Long
    Dim strDevice As String
On Error GoTo errHandle
    
    If mobjFailedImgs Is Nothing Then Exit Sub
    If mobjFailedImgs.Count <= 0 Or mlngFailedLoadCount > 30 Then
        If mobjFailedImgs.Count > 0 Then
            iCols = dcmMiniImage.MultiColumns
            iRows = dcmMiniImage.MultiRows
            
            '如果还有未加载成功的视为下载失败
            For Each strTmpKey In mobjFailedImgs.Keys
                For i = 1 To dcmMiniImage.Images.Count
                    If strTmpKey = dcmMiniImage.Images(i).InstanceUID Then
                        For j = 1 To dcmMiniImage.Images(i).Labels.Count
                            If dcmMiniImage.Images(i).Labels(j).Tag = M_STR_FAILD_TAG Then
                                strDevice = dcmMiniImage.Images(i).Labels(j).Text
                                
                                strDevice = Mid(strDevice, 1, InStr(strDevice, "]"))
                                dcmMiniImage.Images(i).Labels(j).Text = strDevice + "文件下载失败."
                            End If
                        Next
                        
                        dcmMiniImage.Refresh
                        Exit For
                    End If
                Next
            Next
        End If
        
        tmrLoad.Enabled = False
        Exit Sub
    End If
    
    mlngFailedLoadCount = mlngFailedLoadCount + 1
    
    For Each strTmpKey In mobjFailedImgs.Keys
        strTmpFile = mobjFailedImgs(strTmpKey)
        
        '已下载到本地，则替换原来的标记图片
        If Dir(strTmpFile) <> vbNullString Then
            For i = 1 To dcmMiniImage.Images.Count
                If strTmpKey = dcmMiniImage.Images(i).InstanceUID Then
                    Set dcmTag = dcmMiniImage.Images(i).Tag
                                        
                    Set objTmpImg = ReadViewImage(strTmpFile)
                    If err.Number <> 0 Then
                        err.Clear
                        Exit For
                    End If
                    
                    Call dcmMiniImage.Images.Remove(i)
                    
                    Set objTmpImg.Tag = dcmTag '设置图像标记

                    With objTmpImg
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                    
                    '绘制选择框
                    If mblnIsShowCheckbox Then Call DrawItemCheckBorder(objTmpImg)
                    '画报告图标记
                    Call DrawReportImgTag(objTmpImg)
                    
                    Call dcmMiniImage.Images.Add(objTmpImg)
                    
                    Call dcmMiniImage.Images.Move(dcmMiniImage.Images.Count, i)
                    Call mobjFailedImgs.Remove(strTmpKey)
                    
                    Exit For
                End If
            Next
        End If
    Next
    
    Exit Sub
errHandle:
End Sub


Private Sub ucPage_OnPageChange(ByVal lngPageIndex As Long, ByVal lngPageCount As Long)
On Error GoTo errHandle
    Call LoadImage(lngPageIndex, lngPageCount)
    
    '绘制图像的各种标注
    Call DrawImageLabels(dcmMiniImage)
        
    Call UserControl_Resize
Exit Sub
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Sub

'重新下载失败图像
Public Sub ReLoadFailedImage()

    Call dcmMiniImage.Images.Clear
     
    Call LoadImage(1, ucPage.PageRecord)

    '绘制图像的各种标注
    Call DrawImageLabels(dcmMiniImage)
End Sub


Private Function LoadImage(ByVal lngPageIndex As Long, ByVal lngPageCount As Long, Optional ByVal blnGetPath As Boolean) As Boolean
    Dim rsData As ADODB.Recordset

On Error GoTo errHandle
    LoadImage = True
    
    If mstrQueryValue = "0" Then Exit Function
    
    If mslQueryLevel = slLocal Then
        Set rsData = GetImageViewDataFromLocal(lngPageIndex, lngPageCount)
    Else
        If mblnOnlyLoadReportImage Then
            '根据 影像检查记录.报告图像 字段中的值下载，如果为空， 则下载所有报告图像
            Set rsData = GetImageRptData(mstrQueryValue, lngPageIndex, lngPageCount)
            
            mblnIsLoadReportImage = rsData.RecordCount > 0
            
            If rsData.RecordCount <= 0 Then
                Set rsData = GetImageViewData(mslQueryLevel, mstrQueryValue, lngPageIndex, lngPageCount, mblnQueryTmpRecord)
            End If
        Else
            Set rsData = GetImageViewData(mslQueryLevel, mstrQueryValue, lngPageIndex, lngPageCount, mblnQueryTmpRecord)
        End If
    End If
    
    If blnGetPath Then
        Set mrsRecord = rsData
        Exit Function
    End If
    If rsData Is Nothing Then Exit Function
        
    If mslQueryLevel = slLocal Then
        LoadImage = LoadViewImageToFaceFromLocal(rsData, dcmMiniImage)
    Else
        If ImgLoadType = FileLoadType.Normal Then
            LoadImage = LoadViewImageToFaceWithNormal(rsData, dcmMiniImage)     '使用原始模式加载
        Else
            LoadImage = LoadViewImageToFaceWithService(rsData, dcmMiniImage)    '使用ZLPacsServerCenter服务,后台加载
        End If
    End If
    
    Exit Function
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Function

Private Sub UserControl_Initialize()
    
    dcmMiniImage.CellSpacing = 3
    mblnEnable = True
    
    mblnIsShowCheckbox = False
    mblnIsShowPopup = False
    mblnShowPageControl = False
    
    mlngBigImageWay = 0
    
    mstrQueryValue = ""
    mlngSelectIndex = 0
    
    mnuReUpLoad.Visible = False
    mnuReUpLoad.Enabled = False
    
    ucPage.PageRecord = 5
    mblnIsAutoHidePageControl = True
End Sub




Public Sub ClearChecked()
'清除选择
    Dim i As Long
    
    For i = 1 To dcmMiniImage.Images.Count
        ImgChecked(i) = False
    Next i
End Sub



Public Sub SelectedAll()
'全选
    Dim i As Long
    
    For i = 1 To dcmMiniImage.Images.Count
        ImgChecked(i) = True
    Next i
End Sub



Private Sub UserControl_Resize()
    Dim iCols As Integer, iRows As Integer
    Dim i As Long, j As Long
    Dim Img As DicomImage
    Dim sngW As Single '黄框占图像比例
    
On Error Resume Next
    
    dcmMiniImage.Left = 0
    dcmMiniImage.Top = 0
    dcmMiniImage.Width = UserControl.ScaleWidth
    If mblnIsDock Then
        dcmMiniImage.Height = UserControl.ScaleHeight - IIf(mblnShowPageControl, ucPage.Height + 480, 420)
    Else
        dcmMiniImage.Height = UserControl.ScaleHeight - IIf(mblnShowPageControl, ucPage.Height + 60, 0)
    End If
    
    ucPage.Left = 0
    ucPage.Top = dcmMiniImage.Height + 30
    
    ResizeRegion dcmMiniImage.Images.Count, dcmMiniImage.Width, dcmMiniImage.Height, iRows, iCols
    dcmMiniImage.MultiColumns = iCols
    dcmMiniImage.MultiRows = iRows
    
    '判断是否黄框占据图片超过20%
    If dcmMiniImage.Images.Count > 0 Then
        Set Img = dcmMiniImage.Images(mlngSelectIndex)
        sngW = CON_INT_DICOMSELECTWIDTH / (Img.SizeX * Img.ActualZoom)
    End If

    If sngW > 0.2 Then
        '未多选图像并且黄框占据图片超过20%，需要隐藏选中框
        For i = 1 To dcmMiniImage.Images.Count
            For j = 1 To dcmMiniImage.Images(i).Labels.Count
                If dcmMiniImage.Images(i).Labels(j).Tag = M_STR_SELECT_TAG Then dcmMiniImage.Images(i).Labels(j).Visible = False
            Next
        Next
    Else
        '显示选中框
        For i = 1 To dcmMiniImage.Images.Count
            For j = 1 To dcmMiniImage.Images(i).Labels.Count
                If dcmMiniImage.Images(i).Labels(j).Tag = M_STR_SELECT_TAG Then dcmMiniImage.Images(i).Labels(j).Visible = True
            Next
        Next
    End If

    '以前功能保留
    For i = 1 To dcmMiniImage.Images.Count
        If mobjFailedImgs.Exists(dcmMiniImage.Images(i).InstanceUID) Then
            For j = 1 To dcmMiniImage.Images(i).Labels.Count
                If dcmMiniImage.Images(i).Labels(j).Tag = M_STR_FAILD_TAG Then
                    dcmMiniImage.Images(i).Labels(j).Left = 0
                    dcmMiniImage.Images(i).Labels(j).Top = dcmMiniImage.Height / Screen.TwipsPerPixelY / iRows - dcmMiniImage.Images(i).Labels(j).Height * 2
                End If
            Next
        End If
    Next

    err.Clear
End Sub



Private Function GetImageRow(ByVal lngImageIndex As Long) As Integer
'取得当前所在行
    GetImageRow = CInt(lngImageIndex / dcmMiniImage.MultiColumns) + 1
End Function

Private Sub UserControl_Show()
    Call UserControl_Resize
End Sub


Private Sub DrawItemCheckBorder(dcmImg As DicomImage)
    Dim lSelect As DicomLabel
    Dim lBorder As DicomLabel
    
    Set lBorder = New DicomLabel

    With lBorder
        .LabelType = 2            '边框
        .Width = 1000
        .Height = 1000
        .Left = 0
        .Top = 0
        .LineWidth = 2
    
    
        .ForeColour = vbYellow
        .BackColour = vbYellow
    
    
        .Transparent = True
        .ScaleWithCell = True
        .Tag = M_STR_BORDER_TAG
    
        .Visible = True
    End With
    
    dcmImg.Labels.Add lBorder
    


    Set lSelect = New DicomLabel
    
    With lSelect
        .LabelType = 2            '矩形
        .Width = CON_INT_DICOMSELECTWIDTH
        .Height = CON_INT_DICOMSELECTWIDTH
        .Left = 1
        .Top = 1
        .LineWidth = 2
        
        .ForeColour = vbYellow
        .BackColour = vbRed
        
                
        .Transparent = True
        .ScaleWithCell = False
        .ImageTied = False
    
        .Tag = M_STR_SELECT_TAG
        
        .Visible = True
    End With
    
    dcmImg.Labels.Add lSelect
    
    dcmImg.BorderStyle = vbRed
End Sub

Public Sub DrawReportImgTag(dcmImg As DicomImage)
    Dim lRpt As DicomLabel
    Dim i As Integer
    
    
    If dcmImg.Tag.ReportImage <> "" Then
        Set lRpt = New DicomLabel
                
        With lRpt
            .LabelType = doLabelText
            .Width = 300
            .Height = 80
            .ImageTied = False
            .Transparent = True
            .ScaleWithCell = True
            .ScaleFontSize = 40
            .Font.Name = "宋体"
            .Font.Size = 40
            .Font.Bold = True
            .ForeColour = vbWhite
            .BackColour = vbRed
            .Left = 350
            .Top = 20
            .Text = "报告图"
            .ShowTextBox = True
            .Shadow = doShadowBottomRight
            .Alignment = doAlignCentre
            .Visible = True
            .Tag = "报告图"
        End With
        
        dcmImg.Labels.Add lRpt
    Else
        For i = 1 To dcmImg.Labels.Count
            '如果移除了一个标注，标注总数会减少，判断是否已经处理完所有标注，并将i减一
            If i > dcmImg.Labels.Count Then Exit For
            If dcmImg.Labels(i).Tag = "报告图" Then
                Call dcmImg.Labels.Remove(i)
                i = i - 1
            End If
        Next i
    End If
    
    dcmImg.Refresh False
End Sub

Private Sub DrawImageLabels(dcmViewer As DicomViewer)
'绘制图像的各种标注
    Dim i As Long

    '循环每一个图像，画标注
    For i = 1 To dcmViewer.Images.Count
        '画选择框
        If mblnIsShowCheckbox Then
            Call DrawItemCheckBorder(dcmViewer.Images(i))
        End If
        '画报告图标记
        Call DrawReportImgTag(dcmViewer.Images(i))
    Next i
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    
    dcmMiniImage.CellSpacing = PropBag.ReadProperty("CellSpacing", 3)
    dcmMiniImage.BackColour = PropBag.ReadProperty("BackColor", vbBlack)
    mblnEnable = PropBag.ReadProperty("Enable", True)
    mblnIsShowCheckbox = PropBag.ReadProperty("ShowCheckbox", False)
    mblnIsShowPopup = PropBag.ReadProperty("ShowPopup", False)
    ucPage.PageRecord = PropBag.ReadProperty("PageImgCount", 5)
    AutoRedraw = PropBag.ReadProperty("AutoRedrawStyle", False)
    mlngMouseMoveZoom = PropBag.ReadProperty("MouseMoveZoom", 0)
    
    ucPage.AutoRedrawStyle = AutoRedraw
    
    err.Clear
End Sub



Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next

    Call PropBag.WriteProperty("CellSpacing", dcmMiniImage.CellSpacing, 3)
    Call PropBag.WriteProperty("BackColor", dcmMiniImage.BackColour, vbBlack)
    Call PropBag.WriteProperty("Enable", mblnEnable, True)
    Call PropBag.WriteProperty("ShowCheckbox", mblnIsShowCheckbox, False)
    Call PropBag.WriteProperty("ShowPopup", mblnIsShowPopup, False)
    Call PropBag.WriteProperty("PageImgCount", ucPage.PageRecord, 5)
    Call PropBag.WriteProperty("AutoRedrawStyle", AutoRedraw, False)
    Call PropBag.WriteProperty("MouseMoveZoom", mlngMouseMoveZoom, 0)
    
    err.Clear
End Sub

Private Sub ShowPageControl()
'显示分页工具栏
On Error GoTo errHandle
    mblnShowPageControl = True
    ucPage.Visible = mblnShowPageControl
    
    Call UserControl_Resize
errHandle:
End Sub

'临时方法：获取所有图像缓存路径
Public Function GetPathString() As String
    Dim strTmpFile As String

    Call LoadImage(1, ucPage.RecordCount, True)
    
    If mrsRecord Is Nothing Then Exit Function
    If mrsRecord.RecordCount <= 0 Then Exit Function
    
    strTmpFile = ""
    mrsRecord.MoveFirst
    If mslQueryLevel = slLocal Then
        Do While Not mrsRecord.EOF
            strTmpFile = strTmpFile & "|" & Trim(NVL(mrsRecord!路径))
            mrsRecord.MoveNext
        Loop
    Else
        '设置音视频的显示文件，如果为音视频文件时，该过程将不从服务器中直接下载数据文件
        Do While Not mrsRecord.EOF
            If NVL(mrsRecord!动态图, imgTag) <> VIDEOTAG And NVL(mrsRecord!动态图, imgTag) <> AUDIOTAG Then
                strTmpFile = strTmpFile & "|" & GetCacheDir & NVL(mrsRecord("URL")) & IIf(mblnIsLoadReportImage, ".jpg", "")
            End If
            mrsRecord.MoveNext
        Loop
    End If
    
    Set mrsRecord = Nothing
    GetPathString = strTmpFile
End Function

Public Sub AfterSaveStudy(dcmImage As DicomImage)
    If Not mobjImageProcess Is Nothing Then
        mobjImageProcess.AfterSaveStudy dcmImage
    End If
End Sub


