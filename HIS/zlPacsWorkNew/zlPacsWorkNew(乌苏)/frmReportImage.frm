VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{84865D89-6B2D-42E2-98C7-18F4206945F5}#2.0#0"; "zl9PacsControl.ocx"
Begin VB.Form frmReportImage 
   BorderStyle     =   0  'None
   Caption         =   "报告图像"
   ClientHeight    =   6765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picMiniCache 
      Height          =   3855
      Left            =   4080
      ScaleHeight     =   3795
      ScaleWidth      =   4155
      TabIndex        =   18
      Top             =   2760
      Width           =   4215
      Begin VB.ComboBox cboCache 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   480
         Width           =   2415
      End
      Begin zl9PacsControl.ucImagePreview ucMiniCache 
         Height          =   1215
         Left            =   240
         TabIndex        =   19
         Top             =   2400
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   2143
         BackColor       =   -2147483629
         ShowCheckbox    =   -1  'True
      End
   End
   Begin VB.PictureBox picMiniViewer 
      Height          =   1365
      Left            =   240
      ScaleHeight     =   1305
      ScaleWidth      =   3615
      TabIndex        =   16
      Top             =   5280
      Width           =   3675
      Begin zl9PacsControl.ucImagePreview ucMiniImageViewer 
         Height          =   975
         Left            =   45
         TabIndex        =   17
         Top             =   120
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   1720
         BackColor       =   -2147483629
      End
   End
   Begin VB.PictureBox picMenu 
      Height          =   540
      Left            =   2100
      ScaleHeight     =   480
      ScaleWidth      =   525
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   150
      Visible         =   0   'False
      Width           =   585
      Begin XtremeCommandBars.CommandBars cbrMain 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin MSComctlLib.ImageList listCur 
      Bindings        =   "frmReportImage.frx":0000
      Left            =   1200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportImage.frx":0014
            Key             =   "Pen"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMiniImageC 
      Height          =   1455
      Left            =   240
      ScaleHeight     =   1395
      ScaleWidth      =   3675
      TabIndex        =   3
      Top             =   3600
      Width           =   3735
      Begin VB.VScrollBar vscrollMini 
         Height          =   975
         Left            =   3360
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
      Begin DicomObjects.DicomViewer dcmMiniImageC 
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3255
         _Version        =   262147
         _ExtentX        =   5741
         _ExtentY        =   1931
         _StockProps     =   35
         BackColor       =   -2147483629
      End
   End
   Begin VB.PictureBox picReportImage 
      Height          =   2055
      Left            =   3600
      ScaleHeight     =   1995
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      Begin DicomObjects.DicomViewer dcmReportImage 
         Height          =   1695
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
         _Version        =   262147
         _ExtentX        =   3413
         _ExtentY        =   2990
         _StockProps     =   35
         BackColor       =   -2147483629
      End
   End
   Begin VB.PictureBox picMark 
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2595
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   840
      Width           =   3015
      Begin VB.PictureBox picNumMark 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1030
         Left            =   300
         ScaleHeight     =   1035
         ScaleWidth      =   2040
         TabIndex        =   6
         Top             =   1300
         Width           =   2040
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":032E
            Height          =   510
            Index           =   1
            Left            =   490
            Picture         =   "frmReportImage.frx":0F70
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":1BB2
            Height          =   510
            Index           =   4
            Left            =   510
            Picture         =   "frmReportImage.frx":27F4
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   510
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":3436
            Height          =   510
            Index           =   2
            Left            =   1000
            Picture         =   "frmReportImage.frx":4078
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   0
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":4CBA
            Height          =   510
            Index           =   5
            Left            =   1010
            Picture         =   "frmReportImage.frx":58FC
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   510
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":653E
            Height          =   510
            Index           =   3
            Left            =   1560
            Picture         =   "frmReportImage.frx":7180
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   0
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":7DC2
            Height          =   510
            Index           =   6
            Left            =   1510
            Picture         =   "frmReportImage.frx":8A04
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   510
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":9646
            Height          =   1020
            Index           =   0
            Left            =   0
            Picture         =   "frmReportImage.frx":A288
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "自动编号"
            Top             =   0
            Value           =   1  'Checked
            Width           =   510
         End
      End
      Begin DicomObjects.DicomViewer dcmMark 
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1695
         _Version        =   262147
         _ExtentX        =   2990
         _ExtentY        =   1720
         _StockProps     =   35
         BackColor       =   -2147483629
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   120
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmReportImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private mblnSingleWindow As Boolean     '是否使用独立窗口显示报告编辑器，True-独立窗口显示；False-嵌入式显示
Private mlngAdviceID As Long    '医嘱ID
Private mintEditType As Integer '病历状态 0 创建，1书写，2 修订
Private mlngReportID As Long    '报告内容ID
Private mlngFileID As Long      '报告单格式ID
Private mlngShowBigImg As Long          '是否显示大图,0-不显示；1-鼠标移动时显示；2-鼠标单击显示独立窗口
Private mdblBigImgZoom As Double        '报告大图放大倍数
Private mintImageDblClick As Integer    '缩略图双击后的操作 0--直接写入报告；1--打开图片编辑窗口
Private mblnEditable As Boolean         '是否可以编辑内容
Private mintMoustType As Integer        '鼠标工作类型
Private mblnUserInvoke As Boolean       '是否用户操作触发
Private mblnMoved As Boolean            '是否已经转储
Private mintCurImgIndex As Integer      '当前选中的图像
Private mintShowPhotoNumber As Integer  '当前界面能够显示出图像的最佳数量
Private mlngModule As Long

Public mSelMiniImg As DicomImage
Private mSelReportImg As DicomImage
Private mSelViewerIndex As Integer  '当前被选中的报告图象框ID，从1开始计数
Private mselReportImgIndex As Integer   '当前被选中的报告图像ID，从1开始计数
Private mdblMarkZoom As Double          '当前标记图中实际像素和标记之间的缩放比例
Private lngColor(10) As Long             '标记图中圆形编号使用的9个颜色
Private mlngCY1 As Long                 '标记图的高度
Private mlngMarkW As Long               '标记图的宽度
Private mlngCY2 As Long                 '报告图的高度
Private mlngRptImgW As Long             '报告图的宽度
Private mlngCY3 As Long                 '缩略图图的高度

Public pMarkModified As Boolean        '标记图的标记有改动
Public pImageModified As Boolean       '记录报告图像是否修改，如果没有修改，则保存报告的时候不再保存图像
Public pobjMarks As cPicMarks          '当前标记图的标注对象
Public pMarkImageID As Long            '当前标记图在数据库表“电子病历内容”表中的ID
Public pTableID As String              '当前图像所在表格的ID串，用“;”分隔。


Private mintShowMarkImage As Integer   '是否显示标记图   0-隐藏标记图  1-显示标记图
Private mblnIsInitFace As Boolean        '是否已经加载窗体

Private blnLoadImages As Boolean        '记录本次刷新是否加载了图像


Private mdcmGlobal As New DicomGlobal    '定义UIDRoot=1

Private mblnUseActiveVideo As Boolean
Private mrsImageCache As ADODB.Recordset
Private mdcmUID As New DicomGlobal
Private mlngReleationType As Integer    '1--导出，2--导入
Private mlngCurDeptId As Long
Private mlngStudyState As Long
Private mstrTmpQueryValue As String
Private mblnUseAfterCapture As Boolean
Private mblnTmpUseAfterCapture As Boolean

Public Event AfterReleationImage(ByVal lngReleationType As Long)

Private Enum MarkType
    自动编号 = 0: 编号1: 编号2: 编号3: 编号4: 编号5: 编号6
End Enum

Property Get ImageCount() As Long
    If mblnUseActiveVideo Then
'        ImageCount = mobjStudyImage.Images.CurImageCount
        ImageCount = ucMiniImageViewer.CurImageCount
    Else
        ImageCount = dcmMiniImageC.Images.Count
    End If
End Property

Property Get dcmImages() As Object
    If mblnUseActiveVideo Then
'        Set dcmImages = mobjStudyImage.Images.ImgViewer.Images
        Set dcmImages = ucMiniImageViewer.ImgViewer.Images
    Else
        Set dcmImages = dcmMiniImageC.Images
    End If
End Property

Public Sub MovePage(ByVal lngPageType As TMoveType)
'移动缩略图页面
    If mblnUseActiveVideo Then
        ucMiniImageViewer.MovePage (lngPageType)
    End If
End Sub


Public Sub zlRefresh(ByVal lngAdviceId As Long, FileID As Long, _
        ReportID As Long, blnSingleWindow As Boolean, lngShowBigImg As Long, _
        dblBigImgZoom As Double, intImageDblClick As Integer, blnEditable As Boolean, _
        ByVal blnMoved As Boolean, ByVal intMinImageCount As Integer, blnFormIsSelected As Boolean, _
        ByVal lngModule As Long, ByVal lngCurDeptId As Long, ByVal lngStudyState As Long)
    Dim i As Integer
    Dim intShowMarkImage As Integer
    
    mlngCurDeptId = lngCurDeptId
    mlngStudyState = lngStudyState
    mlngAdviceID = lngAdviceId
    mlngFileID = FileID
    mlngReportID = ReportID
    mlngShowBigImg = lngShowBigImg
    mdblBigImgZoom = dblBigImgZoom
    mintImageDblClick = intImageDblClick
    mblnEditable = blnEditable
    mblnMoved = blnMoved
    mintShowPhotoNumber = intMinImageCount
    mlngModule = lngModule
    mblnSingleWindow = blnSingleWindow
    
    intShowMarkImage = DecideMarkImagesVisible    '判断标记图是否可见
    
    mblnUseAfterCapture = Val(GetDeptPara(mlngCurDeptId, "启用后台采集", 1, True)) = 1
    
    ucMiniImageViewer.BigImageWay = lngShowBigImg
    ucMiniImageViewer.MouseMoveZoom = dblBigImgZoom
    ucMiniImageViewer.ShowPopup = False
    
    
    '判断如果是 独立窗口 或者 没有加载过窗体 或者 标记图状态已经改变，则重新加载初始化界面
    If mblnSingleWindow Or (Not mblnIsInitFace Or (mintShowMarkImage <> intShowMarkImage)) Then
        mintShowMarkImage = intShowMarkImage
        Call InitLoaclParas     '读取本机参数
        Call InitFaceScheme     '初始化窗体界面
    Else
        If mblnTmpUseAfterCapture <> mblnUseAfterCapture Then
            Call InitFaceScheme
        End If
    End If
    
    mblnTmpUseAfterCapture = mblnUseAfterCapture
    
    '重新初始化内部参数
    pTableID = ""
    pMarkImageID = 0
    pImageModified = False
    pMarkModified = False
    dcmMark.Images.Clear
    
    If Not (pobjMarks Is Nothing) Then
        For i = 1 To pobjMarks.Count
            pobjMarks.Remove 1
        Next i
    End If
    
    
    '标记本次刷新还没有加载图像
    blnLoadImages = False
    '如果窗体是正在被显示的，则加载图像
    If blnFormIsSelected = True And Me.Visible Then
        '根据需要加载图像
        Call LoadImages
    Else
        Call ClearReportImages
    End If
    
    '设置界面控件是否可以编辑
    picMark.Enabled = mblnEditable
    picReportImage.Enabled = mblnEditable
    picMiniImageC.Enabled = mblnEditable
    picMiniViewer.Enabled = mblnEditable
End Sub

Private Sub ClearReportImages()
    Dim i As Integer
    
    '初始化各个对象
    For i = 1 To dcmReportImage.Count - 1
        Unload dcmReportImage(i)
    Next i
    dcmMark.Images.Clear
End Sub

Public Sub RefPacsPic(Optional ByVal lngEventType As TVideoEventType = vetUpdateImg)
    '读取和显示当前可选报告图像
    If mblnUseAfterCapture And mlngModule <> 1920 Then
        Call LoadMiniCache(lngEventType)
    End If
    
    If lngEventType <> vetAfterUpdateImg Then Call LoadMiniImages
End Sub

Private Sub cboCache_Click()
    Dim strQueryValue As String
    
    If mrsImageCache.RecordCount <= 0 Then Exit Sub
    
    mrsImageCache.MoveFirst
    Do While Not mrsImageCache.EOF
        If "姓名：" & Nvl(mrsImageCache!姓名) & "  检查号：" & Nvl(mrsImageCache!检查号) & "  序列" & Nvl(mrsImageCache!序列号) = cboCache.Text Then
            strQueryValue = Nvl(mrsImageCache!序列UID)
            Exit Do
        End If
        
        mrsImageCache.MoveNext
    Loop
    
    Call ucMiniCache.RefreshImage(slSeries, strQueryValue, mblnMoved, True, True)
    mstrTmpQueryValue = strQueryValue
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer
    
    Select Case control.ID
        Case comMenu_Cap_Process '图像处理
            Call OpenImageProcessWind
        Case conMenu_Cap_DevSet
            If mblnUseAfterCapture And mlngModule <> 1290 Then Call ucMiniCache.ShowPageConfig
            Call ucMiniImageViewer.ShowPageConfig
        Case conMenu_PacsReport_DelImage    '删除图像
            If dcmReportImage(mSelViewerIndex).Images.Count >= mselReportImgIndex And mselReportImgIndex <> 0 Then
                dcmReportImage(mSelViewerIndex).Images.Remove mselReportImgIndex
                Call picReportImage_Resize
                pImageModified = True
            End If
        Case conMenu_PacsReport_MoveUp      '前移图像
            If mselReportImgIndex > 1 And dcmReportImage(mSelViewerIndex).Images.Count >= mselReportImgIndex Then
                dcmReportImage(mSelViewerIndex).Images.Move mselReportImgIndex, mselReportImgIndex - 1
                pImageModified = True
            End If
        Case conMenu_PacsReport_MoveDown    '后移图像
            If mselReportImgIndex > 0 And dcmReportImage(mSelViewerIndex).Images.Count > mselReportImgIndex Then
                dcmReportImage(mSelViewerIndex).Images.Move mselReportImgIndex, mselReportImgIndex + 1
                pImageModified = True
            End If
        Case conMenu_PacsReport_DelMarks    '清除标注
            If dcmMark.Images.Count > 0 Then
                dcmMark.Images(1).Labels.Clear
                dcmMark.Refresh
                For i = 1 To pobjMarks.Count
                    pobjMarks.Remove 1
                Next i
                pMarkModified = True
            End If
        Case conMenu_View_Refresh           '刷新
            '读取和显示当前可选报告图像
            Call LoadMiniImages
        Case conMenu_PacsReport_DelMiniImage    '删除报告图
            
        Case conMenu_PacsReport_SelMiniImage    '提取报告图
            Dim resImages As DicomImages
            
            Set resImages = frmSelectRepImage.ShowMe(Me, mlngAdviceID, mlngShowBigImg, mdblBigImgZoom)
            '把当前图形添加到图象框中
            If resImages.Count > 0 Then
                For i = 1 To resImages.Count
                    dcmReportImage(mSelViewerIndex).Images.Add resImages(i)
                    dcmReportImage(mSelViewerIndex).Images(dcmReportImage(mSelViewerIndex).Images.Count).BorderColour = vbWhite
                Next i
                dcmReportImage(mSelViewerIndex).CurrentIndex = 1
                Call picReportImage_Resize
                pImageModified = True
            End If
            
        Case conMenu_Edit_Import        '导入报告图到缩略图
            mlngReleationType = 2
            Call ReleationImage
            Call RefPacsPic
        
        Case conMenu_File_ExportAll     '导出报告图到缓存图
            mlngReleationType = 1
            Call ReleationImage
            Call RefPacsPic
        
        Case conMenu_Manage_DeleteImage '删除临时图象
            Call DelTempImage
            Call LoadMiniCache
        
        Case conMenu_Manage_RefreshImg  '刷新缓存
            Call LoadMiniCache
    End Select
End Sub

Private Sub DelTempImage()
    Dim rsImageDatas As ADODB.Recordset
    Dim i As Long
    
    '在数据库中查询图像数据
    Set rsImageDatas = GetReleationImageIds()
    
    If rsImageDatas Is Nothing Then
        Call MsgBoxD(Me, "请选择需要删除的检查图像。", vbInformation, Me.Caption)
        Exit Sub
    End If
    
    '当前检查UID在数据库中不存在，则退出本程序
    If rsImageDatas.RecordCount <= 0 Then
        Call MsgBoxD(Me, "请选择需要删除的检查图像。", vbInformation, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "是否确认删除所选图像？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    If DelTempImages(rsImageDatas) Then
        For i = ucMiniCache.CurImageCount To 1
            If ucMiniCache.ImgChecked(i) Then ucMiniCache.DeleteImage (i)
        Next
    End If
End Sub

Private Function DelTempImages(rsImageDatas As ADODB.Recordset) As Boolean
'删除ftp服务器中的文件
    Dim objSrcFtp As New clsFtp
    Dim strCurFtpIp As String, strCurFtpUser As String, strCurFtpPwd As String
    Dim strImageUID As String
    Dim strVirtualPath As String
    
    DelTempImages = False
    If rsImageDatas.RecordCount <= 0 Then Exit Function
    
    strCurFtpIp = ""
    strCurFtpUser = ""
    strCurFtpPwd = ""
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
        strImageUID = Nvl(rsImageDatas!图像UID)
        strVirtualPath = Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url)
            
        If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
            strCurFtpIp = Nvl(rsImageDatas!host)
            strCurFtpUser = Nvl(rsImageDatas!FtpUser)
            strCurFtpPwd = Nvl(rsImageDatas!FtpPwd)
            
            Call objSrcFtp.FuncFtpConnect(strCurFtpIp, strCurFtpUser, strCurFtpPwd)
        End If

        '删除图像文件，当删除失败后，则退出执行
        Call objSrcFtp.FuncDelFile(strVirtualPath, strImageUID)
        
        '删除可能存在的报告图像
        Call objSrcFtp.FuncDelFile(strVirtualPath, strImageUID & ".jpg")
    
        '图像删除成功后，同步删除数据库中的数据
        Call zlDatabase.ExecuteProcedure("ZL_影像检查_删除临时图像(3,'" & strImageUID & "')", Me.Caption)
        
        rsImageDatas.MoveNext
        
        If Not rsImageDatas.EOF Then
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
                Call objSrcFtp.FuncFtpDisConnect
            End If
        End If
    Wend

    objSrcFtp.FuncFtpDisConnect
    
    DelTempImages = True
    
    Exit Function
ErrHandle:
    objSrcFtp.FuncFtpDisConnect
    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Function

Private Function ReleationImage() As Boolean
    Dim strHint As String
    Dim rsImageDatas As ADODB.Recordset
    
    Set rsImageDatas = GetReleationImageIds()
    
    If rsImageDatas Is Nothing Then
        Call MsgBoxD(Me, "请选择需要进行关联的检查图像。", vbInformation, Me.Caption)
        Exit Function
    End If
        
    '当前检查UID在数据库中不存在，则退出本程序
    If rsImageDatas.RecordCount <= 0 Then
        Call MsgBoxD(Me, "请选择需要进行关联的检查图像。", vbInformation, Me.Caption)
        Exit Function
    End If
    
    If mlngReleationType = 2 Then
        '关联图像提示
        strHint = GetReleationHintInfo(mlngAdviceID, rsImageDatas)
        
        If strHint = "" Then
            Call MsgBoxD(Me, "不能查询到需要关联的数据信息，结束关联。", vbOKOnly, Me.Caption)
            Exit Function
        End If
        
        If MsgBoxD(Me, strHint, vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
        
    Else
        '取消关联提示
        If MsgBoxD(Me, "是否确认对所选图像进行取消关联操作？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
    End If

    If mlngReleationType = 2 Then '等于2表示关联图像
        ReleationImage = StartReleation(mlngAdviceID, rsImageDatas)
    Else
        ReleationImage = CancelReleation(mlngAdviceID, rsImageDatas)
    End If
    
    RaiseEvent AfterReleationImage(mlngReleationType)
End Function

'取得关联提示信息
Private Function GetReleationHintInfo(lngAdviceId As Long, rsReleationImage As ADODB.Recordset) As String
    Dim i As Long
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    Dim strResult As String
    Dim strStudyInf As String
    
    GetReleationHintInfo = ""
    
    If rsReleationImage.RecordCount <= 0 Then Exit Function
    
    Call rsReleationImage.MoveFirst
    While Not rsReleationImage.EOF
        strStudyInf = "[" & Nvl(rsReleationImage!姓名) & "(" & Nvl(rsReleationImage!检查号) & ") " & Nvl(rsReleationImage!性别) & " " & Nvl(rsReleationImage!年龄) & "]"
        
        If InStr(strResult, strStudyInf) <= 0 Then
            If strResult <> "" Then strResult = strResult & "+"
        
            strResult = strResult & strStudyInf
        End If
        Call rsReleationImage.MoveNext
    Wend
    
    strSql = "select 检查号,姓名,性别,年龄 from 影像检查记录 where 医嘱ID=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceId)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    
    GetReleationHintInfo = "是否确认将  " & strResult & "  的图像与  [" & Nvl(rsTemp!姓名) & "(" & Nvl(rsTemp!检查号) & ") " & Nvl(rsTemp!性别) & " " & Nvl(rsTemp!年龄) & "]  的检查进行关联操作？"
End Function

Private Function GetReleationImageIds() As ADODB.Recordset
'查询关联或者要取消关联的图像ID
    Dim i As Long, j As Long
    Dim strSql As String
    Dim strValues(0 To 80) As String
    Dim strValue As String
    Dim strUninTable As String
    Dim strFilter As String

    j = 0
    strUninTable = ""
    strFilter = ""
    strValue = ""
    
    '构造查询语句
    If mlngReleationType = 1 Then
        For i = 1 To ucMiniImageViewer.CurImageCount
            If ucMiniImageViewer.ImgChecked(i) Then
                If j > 79 Then
                    strFilter = strFilter & " Or 图像UID ='" & ucMiniImageViewer.ImgViewer.Images(i).InstanceUID & "'"
                Else
                    If zlCommFun.ActualLen(strValue) > 3600 Then
                         strValues(j) = Mid(strValue, 2)
                         strUninTable = strUninTable & " Union ALL  Select  Column_Value as 图像UID From Table(Cast(f_Str2list([" & j + 1 & "]) As zlTools.t_Strlist))  " & vbCrLf
                         
                         strValue = ""
                         j = j + 1
                    End If
                    
                    strValue = strValue & "," & ucMiniImageViewer.ImgViewer.Images(i).InstanceUID
                End If
            End If
        Next
    Else
        For i = 1 To ucMiniCache.CurImageCount
            If ucMiniCache.ImgChecked(i) Then
                If j > 79 Then
                    strFilter = strFilter & " Or 图像UID ='" & ucMiniCache.ImgViewer.Images(i).InstanceUID & "'"
                Else
                    If zlCommFun.ActualLen(strValue) > 3600 Then
                         strValues(j) = Mid(strValue, 2)
                         strUninTable = strUninTable & " Union ALL  Select  Column_Value as 图像UID From Table(Cast(f_Str2list([" & j + 1 & "]) As zlTools.t_Strlist))  " & vbCrLf
                         
                         strValue = ""
                         j = j + 1
                    End If
                    
                    strValue = strValue & "," & ucMiniCache.ImgViewer.Images(i).InstanceUID
                End If
            End If
        Next
    End If
    
    If strValue <> "" Then
        strValues(j) = Mid(strValue, 2)
        strUninTable = strUninTable & " Union ALL  Select  Column_Value as 图像UID From Table(Cast(f_Str2list([" & j + 1 & "]) As zlTools.t_Strlist))  " & vbCrLf
    End If
    
    '如果没有需要查找的图像UID，则返回空数据集
    If strUninTable <> "" Then
        strUninTable = Mid(strUninTable, 11)
    Else
        Set GetReleationImageIds = Nothing
        Exit Function
    End If
    
'    If strFilter <> "" Then strFilter = " and ( " & Mid(strFilter, 4) & ")"
    If strFilter <> "" Then strFilter = strUninTable & " Union All Select 图像UID from [影像图象] where  ( " & Mid(strFilter, 4) & ")"
    
    '根据移动的方向不同，源图有可能在“影像临时记录”或者“影像检查记录”中
    '关联时从临时记录搬移到正常记录，取消关联时从正常记录搬移到临时记录
    strSql = "Select /*+ RULE*/ D.FTP用户名 As FtpUser,D.FTP密码 As FtpPwd, Decode(C.位置一,Null,C.位置二,C.位置一) as 设备号," & _
        "D.IP地址 As Host," & _
        "'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL,A.图像UID, c.姓名,c.性别,c.年龄,c.检查号 " & _
        "From 影像检查图象 A, 影像检查序列 B, 影像检查记录 C,影像设备目录 D,(" & Replace(strUninTable, "[影像图象]", "影像检查图象") & ") E " & _
        "Where Decode(C.位置一,Null,C.位置二,C.位置一)=D.设备号(+)" & _
        "And A.序列UID=B.序列UID and B.检查UID=C.检查UID and A.图像UID = E.图像UID " & _
        "Union All " & _
        "Select /*+ RULE*/ D.FTP用户名 As FtpUser,D.FTP密码 As FtpPwd, Decode(C.位置一,Null,C.位置二,C.位置一) as 设备号," & _
        "D.IP地址 As Host," & _
        "'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL,A.图像UID, c.姓名,c.性别,c.年龄,c.检查号 " & _
        "From 影像临时图象 A,影像临时序列 B, 影像临时记录 C,影像设备目录 D,(" & Replace(strUninTable, "[影像图象]", "影像临时图象") & ") E " & _
        "Where Decode(C.位置一,Null,C.位置二,C.位置一)=D.设备号(+)" & _
        "And A.序列UID=B.序列UID and B.检查UID=C.检查UID and A.图像UID= E.图像UID"
        
    If mblnMoved Then
        strSql = Replace(strSql, "影像检查图象", "H影像检查图象")
        strSql = Replace(strSql, "影像检查序列", "H影像检查序列")
        strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
    End If
    
    Set GetReleationImageIds = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strValues(0), strValues(1), strValues(2), strValues(3), _
        strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10), _
        strValues(11), strValues(12), strValues(13), strValues(14), strValues(15), strValues(16), strValues(17), _
        strValues(18), strValues(19), strValues(20), strValues(21), strValues(22), strValues(23), strValues(24), strValues(25), strValues(26), _
        strValues(27), strValues(28), strValues(29), strValues(30), strValues(31), strValues(32), strValues(33), strValues(34), strValues(35), strValues(36), _
        strValues(37), strValues(38), strValues(39), strValues(40), strValues(41), strValues(42), strValues(43), strValues(44), strValues(45), strValues(46), _
        strValues(47), strValues(48), strValues(49), strValues(50), strValues(51), strValues(52), strValues(53), strValues(54), strValues(55), strValues(56), _
        strValues(57), strValues(58), strValues(59), strValues(60), strValues(61), strValues(62), strValues(63), strValues(64), strValues(65), strValues(66), _
        strValues(67), strValues(68), strValues(69), strValues(70), strValues(71), strValues(72), strValues(73), strValues(74), strValues(75), strValues(76), _
        strValues(77), strValues(78), strValues(79), strValues(80))
End Function

Private Function StartReleation(ByVal lngAdviceId As Long, rsImageDatas As ADODB.Recordset) As Boolean
'开始关联
On Error GoTo ErrHandle
    Dim strSql As String
    Dim strNewStudyUID As String, strNewSeriesUid As String
    Dim curDate As Date
    Dim strReportImageIds As String
    Dim strOldReportImages As String
    Dim lngReportImageLen As Long
    Dim blnBeginTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim rsReportImage As ADODB.Recordset
    Dim objMoveList As New Collection
    Dim strNewDeviceNo As String, strNewFtpIp As String, strNewFtpUrl As String, strNewFtpVirtualPath As String, strNewFtpUser As String, strNewFtpPwd As String
    
    blnBeginTrans = False
    StartReleation = False
    
    curDate = zlDatabase.Currentdate
    
    strSql = "select 检查UID,接收日期 from 影像检查记录 where 医嘱ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceId)
    
    If rsTmp.RecordCount <= 0 Then
        Call MsgBoxD(Me, "找不到待关联的检查信息。", vbInformation, Me.Caption)
        Exit Function
    End If
    
    If Trim(Nvl(rsTmp!检查uid)) = "" Or Trim(Nvl(rsTmp!接收日期)) = "" Then
        
        '尚未采集图像，需要生成新的检查UID
        strNewStudyUID = CreateStudyUid(mdcmUID.NewUID)
        
        Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
        If Trim(strNewFtpIp) = "" Then
            Call MsgBoxD(Me, "不能取得有效的存储设备，请检查存储设备配置。", vbInformation, Me.Caption)
            Exit Function
        End If
        
        '更新存储设备信息
        strSql = "Zl_影像检查_更新设备(" & mlngAdviceID & ",'" & strNewStudyUID & "','" & strNewDeviceNo & "'," & _
                                        "to_Date('" & Format(curDate, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    Else
        strNewStudyUID = Nvl(rsTmp!检查uid)
        
        Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
        If Trim(strNewFtpIp) = "" Then
            Call MsgBoxD(Me, "不能取得有效的存储设备，请检查存储设备配置。", vbInformation, Me.Caption)
            Exit Function
        End If
    End If
    
    '移动图像文件
    If Not MoveImageToStudy(rsImageDatas, strNewStudyUID, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd, objMoveList) Then
        Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
        
        Call MsgBoxD(Me, "图像移动失败，请检查FTP传输是否正常。", vbInformation, Me.Caption)
        Exit Function
    End If
    
    '获取报告图像信息
    strSql = "Select 检查UID,报告图象 From 影像检查记录 Where 医嘱ID=[1]"
    Set rsReportImage = zlDatabase.OpenSQLRecord(strSql, "关联影像", lngAdviceId)
    
    strOldReportImages = ""
    lngReportImageLen = 0
    
    If rsReportImage.RecordCount > 0 Then
        strOldReportImages = Nvl(rsReportImage!报告图象)
        lngReportImageLen = Len(strOldReportImages)
    End If
        
    '创建新的序列UID
    strNewSeriesUid = CreateSeriesUid(mdcmUID.NewUID)
    
    strReportImageIds = ""
    rsImageDatas.MoveFirst
                
    gcnOracle.BeginTrans
    
    blnBeginTrans = True
    
    While Not rsImageDatas.EOF
        '更新图像关联数据
        strSql = "Zl_影像检查_图像关联(" & mlngAdviceID & ",'" & strNewStudyUID & "','" & strNewSeriesUid & "','" & Nvl(rsImageDatas!图像UID) & "')"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        '保存报告数据
        If InStr(1, strOldReportImages & ";" & strReportImageIds, Nvl(rsImageDatas!图像UID)) <= 0 And Len(strReportImageIds) < 4000 - lngReportImageLen - 60 Then
            If strReportImageIds <> "" Then strReportImageIds = strReportImageIds & ";"
            strReportImageIds = strReportImageIds & Nvl(rsImageDatas!图像UID) & ".jpg"
        End If
    
        rsImageDatas.MoveNext
    Wend
    
    '如果需要保持报告图，则需要先查询目前已经保持的报告图像UID
    If rsReportImage.RecordCount > 0 Then
        strReportImageIds = IIf(strOldReportImages <> "", strOldReportImages & ";", "") & strReportImageIds
        strReportImageIds = Replace(strReportImageIds, ";;", ";")
    End If
    
    strSql = "Zl_影像检查_更新报告图(" & mlngAdviceID & ",'" & strReportImageIds & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    '提交事务
    Call gcnOracle.CommitTrans
    
    Call ClearFtpImage(rsImageDatas, strNewStudyUID)
    
    StartReleation = True
    
    Exit Function
ErrHandle:
    If blnBeginTrans Then Call gcnOracle.RollbackTrans
    
    Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
    Call OutputDebug("StartReleation", err)
    
    Call RaiseErr(err)  '继续抛出错误
End Function

Private Function CancelReleation(ByVal lngAdviceId As Long, rsImageDatas As ADODB.Recordset) As Boolean
'撤销关联
On Error GoTo ErrHandle
    Dim strSql As String
    Dim strNewStudyUID As String, strNewSeriesUid As String
    Dim curDate As Date
    Dim strReportImageIds As String
    Dim strOldReportImages As Long
    Dim lngReportImageLen As Long
    Dim blnBeginTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim rsReportImage As ADODB.Recordset
    Dim objMoveList As New Collection
    Dim strNewDeviceNo As String, strNewFtpIp As String, strNewFtpUrl As String, strNewFtpVirtualPath As String, strNewFtpUser As String, strNewFtpPwd As String
    
    blnBeginTrans = False
    CancelReleation = False
    
    curDate = zlDatabase.Currentdate
    
    '撤销图像关联
    strNewStudyUID = CreateStudyUid(mdcmUID.NewUID)
        
    Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
    If Trim(strNewFtpIp) = "" Then
        Call MsgBoxD(Me, "不能取得有效的存储设备，请检查存储设备配置。", vbInformation, Me.Caption)
        Exit Function
    End If
    
    '移动图像文件
    If Not MoveImageToStudy(rsImageDatas, strNewStudyUID, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd, objMoveList) Then
        Call MsgBoxD(Me, "图像移动失败，请检查FTP传输是否正常。", vbInformation, Me.Caption)
        Exit Function
    End If
    
    strSql = "Select 检查UID,报告图象 From 影像检查记录 Where 医嘱ID=[1]"
    Set rsReportImage = zlDatabase.OpenSQLRecord(strSql, "关联影像", mlngAdviceID)
    
    If rsReportImage.RecordCount > 0 Then
        strReportImageIds = Nvl(rsReportImage!报告图象)
        strReportImageIds = Replace(strReportImageIds, " ", "") '采集图像时，可能会在报告图数据后添加空格
    End If
    
    '更新数据
    rsImageDatas.MoveFirst
    
    gcnOracle.BeginTrans
    
    blnBeginTrans = True
    
    While Not rsImageDatas.EOF
        strSql = "Select D.检查UID From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像临时序列 D " & _
                 "Where C.医嘱ID=[1] And A.图像UID=[2] And A.序列UID=B.序列UID And B.检查UID=C.检查UID And A.序列UID = D.序列UID"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "关联影像", mlngAdviceID, Nvl(rsImageDatas!图像UID))

        If rsTmp.RecordCount > 0 Then strNewStudyUID = Nvl(rsTmp!检查uid)
            
        strSql = "Zl_影像检查_撤销关联(" & mlngAdviceID & ",'" & Nvl(rsImageDatas!图像UID) & "','" & strNewStudyUID & "','" & strNewDeviceNo & "'," & _
                                        "to_Date('" & Format(curDate, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'))"
                                        
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        '修改报告图数据
        strReportImageIds = Replace(strReportImageIds, Nvl(rsImageDatas!图像UID) & ".jpg;", "")
        strReportImageIds = Replace(strReportImageIds, Nvl(rsImageDatas!图像UID) & ".jpg", "")
        
        rsImageDatas.MoveNext
    Wend
    
    '更新报告图像
    strSql = "Zl_影像检查_更新报告图(" & mlngAdviceID & ",'" & strReportImageIds & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call gcnOracle.CommitTrans
    
    Call ClearFtpImage(rsImageDatas, strNewStudyUID)
    
    CancelReleation = True
Exit Function
ErrHandle:
    If blnBeginTrans Then Call gcnOracle.RollbackTrans
    Call OutputDebug("CancelReleation", err)
    Call RaiseErr(err)
End Function

Private Sub ClearFtpImage(rsImageDatas As ADODB.Recordset, ByVal strNewStudyUID As String)
On Error GoTo ErrHandle
'转移图像成功后，在删除临时图像和原有FTP的图像和目录，清场操作出现错误可以不处理
    Dim objSrcFtp As New clsFtp
    Dim strTmpFile As String
    Dim strVirtualPath As String
    Dim strImageUID As String
    Dim strCurFtpIp As String, strCurFtpUser As String, strCurFtpPwd As String
    Dim strNewDirectory
    
    strCurFtpIp = ""
    strCurFtpUser = ""
    strCurFtpPwd = ""
    strNewDirectory = App.Path & "\TmpImage\" & Format(zlDatabase.Currentdate, "YYYYMMDD")
    
    If Not DirExists(strNewDirectory) Then MkDir strNewDirectory
    If Not DirExists(strNewDirectory & "\" & strNewStudyUID) Then MkDir strNewDirectory & "\" & strNewStudyUID
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
        strTmpFile = App.Path & "\TmpImage\" & Nvl(rsImageDatas!图像UID)
        
        strImageUID = Nvl(rsImageDatas!图像UID)
        
        strVirtualPath = Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url)
                
        If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
            strCurFtpIp = Nvl(rsImageDatas!host)
            strCurFtpUser = Nvl(rsImageDatas!FtpUser)
            strCurFtpPwd = Nvl(rsImageDatas!FtpPwd)
            
            Call objSrcFtp.FuncFtpConnect(strCurFtpIp, strCurFtpUser, strCurFtpPwd)
        End If
        
'       为避免重新下载图像，如果本地存在图像文件，则不用进行删除
        
        If FileExists(strTmpFile) Then Call Kill(strTmpFile)
        If FileExists(strTmpFile & ".jpg") Then Call Kill(strTmpFile & ".jpg")

        '移动文件到新的位置
        Call MoveFile(App.Path & "\TmpImage\" & Nvl(rsImageDatas!Url) & "\" & Nvl(rsImageDatas!图像UID), _
            strNewDirectory & "\" & strNewStudyUID & "\" & Nvl(rsImageDatas!图像UID))
        
        Call objSrcFtp.FuncDelFile(strVirtualPath, strImageUID)
        
        '删除空的ftp目录
        Call objSrcFtp.FuncFtpDelDir(Replace(strVirtualPath, strImageUID, ""), strImageUID)
                
        rsImageDatas.MoveNext
        
        If Not rsImageDatas.EOF Then
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
                Call objSrcFtp.FuncFtpDisConnect
            End If
        End If
    Wend
    
    objSrcFtp.FuncFtpDisConnect
Exit Sub
ErrHandle:
    Call OutputDebug("ClearFtpImage", err)
End Sub

'撤销图像的移动
Private Sub CancelImageMove(ByVal strFTPIP As String, ByVal strFTPUser As String, ByVal strFTPPwd As String, objMoveList As Collection)
    Dim i As Long
    Dim objFtp As New clsFtp
    Dim strDestFile As String
    Dim strMoveFile As String
    
    If objMoveList Is Nothing Then Exit Sub
    If objMoveList.Count <= 0 Then Exit Sub
    
On Error GoTo ErrHandle

    Call objFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
    
    For i = 1 To objMoveList.Count
        strDestFile = objMoveList.Item(i)
        
        strMoveFile = Mid(strDestFile, InStr(strDestFile, ">") + 1, 255)
        strDestFile = Mid(strDestFile, 1, InStr(strDestFile, ">") - 1)
        
        Call objFtp.FuncReNameFile(strMoveFile, strDestFile)
    Next i
        
ErrHandle:
    objFtp.FuncFtpDisConnect
End Sub

Public Function MoveImageToStudy(rsImageDatas As ADODB.Recordset, ByVal strNewStudyUID As String, _
    ByVal strFTPIP As String, ByVal strFtpUrl As String, ByVal strFtpVirtualPath As String, _
    ByVal strFTPUser As String, ByVal strFTPPwd As String, ByRef objMoveList As Collection) As Boolean
'------------------------------------------------
'功能：将选定的检查图像移动到ftp上指定的检查中
'返回：True--成功；False－失败
'------------------------------------------------
    Dim objSrcFtp As New clsFtp
    Dim objDestFtp As New clsFtp
    Dim strVirtualPath As String
    Dim strDestPath As String
    Dim strTmpFile As String
    Dim aFiles() As String
    Dim i As Integer
    Dim objFile As New Scripting.FileSystemObject
    Dim lngResult As Long       '记录FTP操作的结果
    Dim strImageUID As String
    Dim dcmImages As New DicomImages
    Dim dcmImg As New DicomImage
    Dim strCurFtpIp As String, strCurFtpUser As String, strCurFtpPwd As String
    Dim strFileList As String
    Dim blnIsMove As Boolean
    Dim strNewDeviceNo As String, strNewFtpIp As String, strNewFtpUrl As String, strNewFtpVirtualPath As String, strNewFtpUser As String, strNewFtpPwd As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo ErrHandle
    
    blnIsMove = False
    MoveImageToStudy = False
    If rsImageDatas.RecordCount <= 0 Then Exit Function

    '连接目标Ftp
    Call objDestFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
    
    strCurFtpIp = ""
    strCurFtpUser = ""
    strCurFtpPwd = ""
    strVirtualPath = ""
    strFileList = ""
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
        strImageUID = Nvl(rsImageDatas!图像UID)
        
        '取消关联
        If mlngReleationType = 1 Then
            strSql = "Select D.检查UID From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像临时序列 D " & _
                     "Where C.医嘱ID=[1] And A.图像UID=[2] And A.序列UID=B.序列UID And B.检查UID=C.检查UID And A.序列UID = D.序列UID"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "关联影像", mlngAdviceID, strImageUID)
    
            If rsTemp.RecordCount > 0 Then strNewStudyUID = Nvl(rsTemp!检查uid)
            
            Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
        End If
        
        If strVirtualPath <> Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url) Then
            strVirtualPath = Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url)
            strFileList = ""
        End If
        
        '当移动的文件不是相同的ftp地址时，则使用下载后再上传的方式转移文件
        If Nvl(rsImageDatas!host) <> strFTPIP Or Nvl(rsImageDatas!Root) <> strFtpUrl Then
        
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
            
                strCurFtpIp = Nvl(rsImageDatas!host)
                strCurFtpUser = Nvl(rsImageDatas!FtpUser)
                strCurFtpPwd = Nvl(rsImageDatas!FtpPwd)
                
                Call objSrcFtp.FuncFtpConnect(strCurFtpIp, strCurFtpUser, strCurFtpPwd)
            End If
        
            strTmpFile = App.Path & "\TmpImage\" & strImageUID
            
            If strFileList = "" Then
                strFileList = objSrcFtp.FuncDirFiles(strVirtualPath)
            End If
            
            '如果源ftp设备中不存在该图像，则不进行移动
            If InStr(strFileList, strImageUID) > 0 Then
                lngResult = objSrcFtp.FuncDownloadFile(strVirtualPath, strTmpFile, strImageUID)
                If lngResult <> 0 Then
                    objSrcFtp.FuncFtpDisConnect
                    objDestFtp.FuncFtpDisConnect
        
                    Call err.Raise(-1, "MoveImageToStudy", "下载关联图像失败。 [图像UID:" & strImageUID & " 文件虚拟目录:" & strVirtualPath & " 本地路径:" & strTmpFile & "]", err.HelpFile, err.HelpContext)
                    Exit Function
                End If
        
                lngResult = objDestFtp.FuncUploadFile(strFtpVirtualPath, strTmpFile, strImageUID)
                If lngResult <> 0 Then
                    objSrcFtp.FuncFtpDisConnect
                    objDestFtp.FuncFtpDisConnect
        
                    Call err.Raise(-1, "MoveImageToStudy", "上传关联图像失败。 [图像UID:" & strImageUID & " 上传虚拟目录:" & strFtpVirtualPath & " 本地路径:" & strTmpFile & "]", err.HelpFile, err.HelpContext)
                    Exit Function
                End If
                
                blnIsMove = True
            End If
        Else
            If strFileList = "" Then
                strFileList = objDestFtp.FuncDirFiles(strVirtualPath)
            End If
            
            '如果源ftp设备中不存在该图像，则不进行移动
            If InStr(strFileList, strImageUID) > 0 Then
                lngResult = objDestFtp.FuncReNameFile(strVirtualPath & "/" & strImageUID, strFtpVirtualPath & "/" & strImageUID)
                If lngResult <> 0 Then
                    '如果文件移动失败，则端开连接重试一次
                    Call objDestFtp.FuncFtpDisConnect
                    Call objDestFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
                    
                    lngResult = objDestFtp.FuncReNameFile(strVirtualPath & "/" & strImageUID, strFtpVirtualPath & "/" & strImageUID)
                    
                    If lngResult <> 0 Then
                        If mlngReleationType = 1 Then Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
                        
                        objSrcFtp.FuncFtpDisConnect
                        objDestFtp.FuncFtpDisConnect
                        
                        Call err.Raise(-1, "MoveImageToStudy", "在Ftp中移动文件时失败。 [图像UID:" & strImageUID & " 原虚拟目录:" & strVirtualPath & " 新虚拟目录:" & strFtpVirtualPath & "]", err.HelpFile, err.HelpContext)
                        Exit Function
                    End If
                End If
                
                blnIsMove = True
                
                '记录已经被移动过的文件，以便在处理数据失败的时候，还可对移动的图像进行恢复
                If Not objMoveList Is Nothing Then
                    Call objMoveList.Add(strVirtualPath & "/" & strImageUID & ">" & strFtpVirtualPath & "/" & strImageUID)
                End If
            End If
        End If
        
        If Nvl(rsImageDatas!host) <> strFTPIP Or Nvl(rsImageDatas!Root) <> strFtpUrl Then
            Call MoveReportImage(strTmpFile, strImageUID, objSrcFtp, objDestFtp, strVirtualPath, strFtpVirtualPath, objMoveList, 0)
        Else
            Call MoveReportImage(strTmpFile, strImageUID, objSrcFtp, objDestFtp, strVirtualPath, strFtpVirtualPath, objMoveList, 1)
            End If
        
        rsImageDatas.MoveNext
        
        If Not rsImageDatas.EOF Then
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
                Call objSrcFtp.FuncFtpDisConnect
            End If
        End If
    Wend
    
    objSrcFtp.FuncFtpDisConnect
    objDestFtp.FuncFtpDisConnect
    
    '如果一个文件都没有被移动，则直接退出
    If Not blnIsMove Then Exit Function
    
    MoveImageToStudy = True
    
    Exit Function
ErrHandle:
    objSrcFtp.FuncFtpDisConnect
    objDestFtp.FuncFtpDisConnect

    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Function

Private Sub MoveReportImage(ByVal strDicomFile As String, ByVal strImgUid As String, _
    objSrcFtp As clsFtp, objDestFtp As clsFtp, ByVal strSourceVirtualPath As String, ByVal strDestVirtualPath As String, _
    objMoveList As Collection, Optional ByVal lngWay As Long = 0)
On Error GoTo ErrHandle
'移动报告图
    Dim dcmImages As New DicomImages
    Dim dcmImg As New DicomImage
    Dim lngResult As Long
    
    If lngWay = 0 Then
        Call objSrcFtp.FuncDelFile(strSourceVirtualPath, strImgUid & ".jpg")
        
        '如果本地中存在从源ftp中下载的dicom图像，则将图像转换成jpg，并保存到目的ftp设备中
        If FileExists(strDicomFile) Then
            Call dcmImages.Clear
            Set dcmImg = dcmImages.ReadFile(strDicomFile)
    
            Call dcmImg.FileExport(strDicomFile & ".jpg", "JPG")
            Call objDestFtp.FuncUploadFile(strDestVirtualPath, strDicomFile & ".jpg", strImgUid & ".jpg")
            
            If FileExists(strDicomFile & ".jpg") Then Call Kill(strDicomFile & ".jpg")
        End If
    Else
        '如果源ftp设备中不存在该图像，则不进行移动
        If objDestFtp.FuncFtpFileExists(strSourceVirtualPath, strImgUid & ".jpg") Then
            lngResult = objDestFtp.FuncReNameFile(strSourceVirtualPath & "/" & strImgUid & ".jpg", strDestVirtualPath & "/" & strImgUid & ".jpg")
            
            If lngResult <> 0 Then
                '如果文件移动失败，则端开连接重试一次
                Call objDestFtp.FuncFtpDisConnect
'                Call objDestFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
                Call objDestFtp.ResotreFtpConnect
                
                Call objDestFtp.FuncReNameFile(strSourceVirtualPath & "/" & strImgUid & ".jpg", strDestVirtualPath & "/" & strImgUid & ".jpg")
                
                '记录已经被移动过的文件，以便在处理数据失败的时候，还可对移动的图像进行恢复
                If Not objMoveList Is Nothing Then
                    Call objMoveList.Add(strSourceVirtualPath & "/" & strImgUid & ".jpg" & ">" & strDestVirtualPath & "/" & strImgUid & ".jpg")
                End If
            End If
        End If
    End If
Exit Sub
ErrHandle:
    Call OutputDebug("MoveReportImage", err)
End Sub

Private Sub GetStorageDevice(ByVal lngAdviceId As Long, ByVal strNewStudyUID As String, _
    ByRef strDeviceNO As String, ByRef strFTPIP As String, _
    ByRef strFtpUrl As String, ByRef strVirtualPath As String, _
    ByRef strFTPUser As String, ByRef strFTPPwd As String)
'获取新的存储设备信息，如果设备存储信息部存在，则需要进行增加
'如果是取消关联，则使用strNewStudyUID将不能从数据库中查找到对应的数据
'strDeviceNum:设备号
'strFtpIp: ftp地址
'strFtpUrl: ftp目录
'strVirtualPath: ftp虚拟存储路径
'strFtpUser: ftp用户名
'strFtpPwd: ftp密码



    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim blnIsGetNewDevice As Boolean
    Dim objDestFtp As New clsFtp
    Dim curDate As Date
    
    strFTPIP = ""
    strFtpUrl = ""
    strFTPUser = ""
    strFTPPwd = ""
    strDeviceNO = ""
    
    strSql = "Select D.FTP用户名 As FtpUser,D.FTP密码 As FtpPwd,C.位置一,C.位置二,C.位置三,C.接收日期," & _
        "D.IP地址 As Host," & _
        "'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL " & _
        "From 影像检查记录 C,影像设备目录 D " & _
        "Where Decode(C.位置一,Null,C.位置二,C.位置一)=D.设备号(+)" & _
        "And C.检查UID= [1] Union All " & _
        "Select D.FTP用户名 As FtpUser,D.FTP密码 As FtpPwd,C.位置一,C.位置二,C.位置三,C.接收日期," & _
        "D.IP地址 As Host," & _
        "'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL " & _
        "From 影像临时记录 C,影像设备目录 D " & _
        "Where Decode(C.位置一,Null,C.位置二,C.位置一)=D.设备号(+)" & _
        "And C.检查UID= [1]"
        
    If mblnMoved Then
        strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
    End If
        
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNewStudyUID)
    
    blnIsGetNewDevice = False
    
    If rsData.RecordCount <= 0 Then
        blnIsGetNewDevice = True
    Else
        '如果执行到这里，说明是执行图像关联,需要判断当前检查的存储设备是否有效，如果无效需生成新的存储设备
        If Trim(rsData!接收日期) = "" Then
            blnIsGetNewDevice = True
        Else
            strDeviceNO = Nvl(rsData!位置一)
            strFTPIP = Nvl(rsData!host)
            strFtpUrl = Nvl(rsData!Root)
            strFTPUser = Nvl(rsData!FtpUser)
            strFTPPwd = Nvl(rsData!FtpPwd)
            strVirtualPath = strFtpUrl & Nvl(rsData!Url)
        End If
    End If
    
    
    If blnIsGetNewDevice Then
        '生成新的检查UID和存储设备,如果执行到这里，说明是取消关联
        
        If mlngModule = 1290 Then
            '查询医技工作站中，检查所对应的存储设备
            strSql = "select d.参数值 " & _
                        " from 医技执行房间 a, 病人医嘱发送 b, 影像DICOM服务对 c, 影像DICOM服务参数 d " & _
                        " Where a.科室ID = b.执行部门id And a.执行间 = b.执行间 And a.检查设备 = c.设备号 " & _
                        " and c.服务功能='图像接收' and c.服务ID=d.服务ID and d.参数名称='存储设备' and b.医嘱id=[1]"
                        
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceId)
            
            If rsTemp.RecordCount <= 0 Then
                MsgBoxD Me, "未找到图像存储设备,请确认当前检查所用设备是否在影像设备目录的服务配置中配置了图像存储。", vbInformation, gstrSysName
                Exit Sub
            End If
            
            strDeviceNO = Nvl(rsTemp!参数值)
        Else
            '查询非医技工作站中的图像存储设备
            strDeviceNO = GetDeptPara(mlngCurDeptId, "存储设备号")
            
            If Val(strDeviceNO) <= 0 Then
                MsgBoxD Me, "未找到图像存储设备,请确认在影像流程管理中是否对该科室配置了图像采集存储设备。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        strSql = "Select 设备号,设备名,'/'||Decode(Ftp目录,Null,'',Ftp目录||'/') As URL,FTP用户名,FTP密码,IP地址 " & _
                    " From 影像设备目录 Where 类型=1 and 设备号=[1] and NVL(状态,0)=1"
                    
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.tag, strDeviceNO)
        
        '如果存储设备停用，则直接退出
        If rsTemp.RecordCount <= 0 Then
            MsgBoxD Me, "未找到存储设备,请确认设备号为 [" & strDeviceNO & "] 的设备是否启用。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strFtpUrl = Nvl(rsTemp("URL"))
        strFTPIP = Nvl(rsTemp("IP地址"))
        strFTPUser = Nvl(rsTemp("FTP用户名"))
        strFTPPwd = Nvl(rsTemp("FTP密码"))
        
        strFtpUrl = IIf(strFtpUrl = "/", "//", strFtpUrl)
        
        objDestFtp.FuncFtpConnect strFTPIP, strFTPUser, strFTPPwd
        On Error GoTo ErrHandle
            curDate = zlDatabase.Currentdate
            
            strVirtualPath = strFtpUrl & Format(curDate, "YYYYMMDD") & "/" & strNewStudyUID
            '创建FTP目录
            objDestFtp.FuncFtpMkDir strFtpUrl, Format(curDate, "YYYYMMDD") & "/" & strNewStudyUID
            
        Call objDestFtp.FuncFtpDisConnect
ErrHandle:
        Call objDestFtp.FuncFtpDisConnect
    End If
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo ErrHandle
    
    Select Case control.ID
        Case conMenu_Edit_Import
            If mlngAdviceID <= 0 Or mlngStudyState < 2 Or mlngStudyState > 4 Then control.Enabled = False
    End Select
    
    Exit Sub
ErrHandle:
    
End Sub

Private Sub chkMark_Click(Index As Integer)
    Dim i As Integer
    If mblnUserInvoke = False Then
        mblnUserInvoke = True
    Select Case Index
        Case 0
            mintMoustType = MarkType.自动编号
        Case 1
            mintMoustType = MarkType.编号1
        Case 2
            mintMoustType = MarkType.编号2
        Case 3
            mintMoustType = MarkType.编号3
        Case 4
            mintMoustType = MarkType.编号4
        Case 5
            mintMoustType = MarkType.编号5
        Case 6
            mintMoustType = MarkType.编号6
    End Select
    For i = 0 To 6
        chkMark(i).value = 0
    Next i
    chkMark(Index).value = 1
    mblnUserInvoke = False
    End If
End Sub

Private Sub dcmMark_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim lTemp As DicomLabel
    Dim strNum As Integer
    
    If Button = 1 And dcmMark.Images.Count > 0 And picMark.MousePointer = 99 Then
        '画标注
        '两种类型的标注，一种是直接自动编号，另一种是手工编号
        pobjMarks.Add pobjMarks.Count + 1
        pobjMarks(pobjMarks.Count).Selected = False
        pobjMarks(pobjMarks.Count).类型 = 6     '圆形编号
        If mintMoustType = MarkType.自动编号 Then
            pobjMarks(pobjMarks.Count).内容 = pobjMarks.Count
        Else
            Select Case mintMoustType
                Case MarkType.编号1
                    pobjMarks(pobjMarks.Count).内容 = 1
                Case MarkType.编号2
                    pobjMarks(pobjMarks.Count).内容 = 2
                Case MarkType.编号3
                    pobjMarks(pobjMarks.Count).内容 = 3
                Case MarkType.编号4
                    pobjMarks(pobjMarks.Count).内容 = 4
                Case MarkType.编号5
                    pobjMarks(pobjMarks.Count).内容 = 5
                Case MarkType.编号6
                    pobjMarks(pobjMarks.Count).内容 = 6
            End Select
        End If
        '点集没有留空
        Set lTemp = New DicomLabel
        lTemp.Left = X
        lTemp.Top = Y
        lTemp.Width = 20
        lTemp.Height = 20
        lTemp.ImageTied = True
        lTemp.Rescale dcmMark.Images(1)
        pobjMarks(pobjMarks.Count).X1 = lTemp.Left / mdblMarkZoom
        pobjMarks(pobjMarks.Count).Y1 = lTemp.Top / mdblMarkZoom
        pobjMarks(pobjMarks.Count).X2 = pobjMarks(pobjMarks.Count).X1
        pobjMarks(pobjMarks.Count).Y2 = pobjMarks(pobjMarks.Count).Y1
        pobjMarks(pobjMarks.Count).填充色 = lngColor(pobjMarks.Count Mod 9 + 1)
        pobjMarks(pobjMarks.Count).填充方式 = -2
        '线条色留空，字体色留空
        pobjMarks(pobjMarks.Count).线型 = 1
        pobjMarks(pobjMarks.Count).线宽 = 1
        Set pobjMarks(pobjMarks.Count).字体 = New StdFont '  "宋体"
        drawPicMarks dcmMark.Images(1), pobjMarks
        dcmMark.Refresh
        
        pMarkModified = True
    End If
End Sub

Private Sub dcmMark_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If dcmMark.Images.Count = 1 Then
        '设置鼠标
        If dcmMark.ImageXPosition(X, Y) > 0 And dcmMark.ImageXPosition(X, Y) < dcmMark.Images(1).SizeX _
           And dcmMark.ImageYPosition(X, Y) > 0 And dcmMark.ImageYPosition(X, Y) < dcmMark.Images(1).SizeY Then
            picMark.MousePointer = 99
            picMark.MouseIcon = listCur.ListImages("Pen").Picture
        Else
            picMark.MousePointer = 0
            picMark.MouseIcon = Nothing
        End If
    End If
End Sub

Private Sub dcmMark_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button = 2 Then ShowPopupMark
End Sub

Private Sub OpenImageProcessWind()
    Call frmReportImageEdit.zlShowMe(mSelMiniImg, Me, mintCurImgIndex, mSelViewerIndex, mlngModule)
End Sub

Private Sub dcmMiniImageC_DblClick()
    
    If dcmMiniImageC.Images.Count > 0 And mSelViewerIndex <> 0 And dcmReportImage.Count > 1 Then
        '判断当前双击的操作动作
        If mintImageDblClick = 0 Then   '直接写入报告
            Dim dcmImage As DicomImage
            Set dcmImage = mSelMiniImg
            
            '调用将当前图形添加到图象框过程
            Call DcmAddImage(dcmImage, mSelViewerIndex)
            
        Else                            '先打开图片编辑窗口
            '先关闭大图窗口
            ReleaseCapture      '解锁鼠标
            frmShowImg.HideMe
            
            Call OpenImageProcessWind
        End If

    End If
End Sub

Public Sub DcmAddImage(dcmImage As DicomImage, SelViewerIndex As Integer)
'把当前图形添加到图象框中
    If Not dcmImage Is Nothing Then
        dcmReportImage(SelViewerIndex).Images.Add dcmImage
        dcmReportImage(SelViewerIndex).Images(dcmReportImage(SelViewerIndex).Images.Count).BorderColour = vbWhite
        dcmReportImage(SelViewerIndex).Images(dcmReportImage(SelViewerIndex).Images.Count).tag = mdcmGlobal.NewUID & ".jpg"
        dcmReportImage(SelViewerIndex).CurrentIndex = 1
        Call picReportImage_Resize
        pImageModified = True
    End If
End Sub

Private Sub dcmMiniImageC_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim i As Integer
    
    If dcmMiniImageC.Images.Count > 0 Then
        For i = 1 To dcmMiniImageC.Images.Count
            dcmMiniImageC.Images(i).BorderColour = vbWhite
            
        Next i
        
        i = dcmMiniImageC.ImageIndex(X, Y)
        If i = 0 Then
            Set mSelMiniImg = dcmMiniImageC.Images(1)
        Else
            Set mSelMiniImg = dcmMiniImageC.Images(i)
        End If
        
        mSelMiniImg.BorderColour = vbRed
        
        mintCurImgIndex = i
        
        '判断是否需要显示大图
        If mlngShowBigImg = 2 Then
            frmShowImg.ShowMe mSelMiniImg, Me, 2, 0, 0
        End If
    End If
End Sub

Private Sub dcmMiniImageC_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim blnShowImg As Boolean
    Dim intCurrImg As Integer

    If dcmMiniImageC.Images.Count <= 0 Or mlngShowBigImg <> 1 Then Exit Sub
    
    '判断是否需要显示图像
    If (0 <= X * Screen.TwipsPerPixelX) And (X * Screen.TwipsPerPixelX <= dcmMiniImageC.Width) And _
       (0 <= Y * Screen.TwipsPerPixelY) And (Y * Screen.TwipsPerPixelY <= dcmMiniImageC.Height) Then
        blnShowImg = True
    End If
    If blnShowImg Then      '显示图像
        SetCapture dcmMiniImageC.hWnd    '锁定鼠标
        
        intCurrImg = dcmMiniImageC.ImageIndex(X, Y)
        If intCurrImg <> 0 Then
            '加载图像并显示
            frmShowImg.ShowMe dcmMiniImageC.Images(intCurrImg), Me, 1, 0, 0, mdblBigImgZoom
        Else
            frmShowImg.HideMe
        End If
    Else        '关闭图像显示
        ReleaseCapture      '解锁鼠标
        frmShowImg.HideMe
    End If
End Sub

Private Sub dcmMiniImageC_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mlngShowBigImg = 1 Then  '关闭大图显示
        ReleaseCapture      '解锁鼠标
        frmShowImg.HideMe
    End If
    
    If Button = 2 Then Call ShowPopupImage(1)
End Sub

Private Sub dcmReportImage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim i As Integer
    
'    If dcmReportImage(Index).Images.Count = 0 Then Exit Sub
    
    mSelViewerIndex = Index
    mselReportImgIndex = dcmReportImage(Index).ImageIndex(X, Y)
    
    For i = 1 To dcmReportImage.Count - 1
        dcmReportImage(i).Labels(1).ForeColour = vbWhite
        dcmReportImage(i).Refresh
    Next i
    dcmReportImage(Index).Labels(1).ForeColour = vbRed
    dcmReportImage(Index).Refresh
    
    If mselReportImgIndex <> 0 Then
        For i = 1 To dcmReportImage(Index).Images.Count
            dcmReportImage(Index).Images(i).BorderColour = vbWhite
        Next i
        dcmReportImage(Index).Images(mselReportImgIndex).BorderColour = vbBlue
    End If
    
    
End Sub

Private Sub dcmReportImage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Long, Y As Long)
    
    If Button = 2 Then Call ShowPopupImage(0)
End Sub

Private Sub ShowPopupCache()

End Sub

Private Sub ShowPopupImage(ByVal intType As Integer)
'------------------------------------------------
'功能：创建鼠标右键弹出菜单
'intType:0--报告图，1--缩略图，2--缓存图
'------------------------------------------------
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrToolPopup As CommandBarPopup
    
    If intType <> 2 Then
        If mblnUseActiveVideo Then
            If ucMiniImageViewer.CurImageCount < 1 Then Exit Sub
        Else
            '如果缩略图没有图像，则禁止右键弹出
            If Me.dcmMiniImageC.Images.Count < 1 Then Exit Sub
        End If
    End If
    
    '鼠标右键弹出菜单
    Set cbrToolBar = cbrMain.Add("鼠标右键", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        If intType = 0 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_DelImage, "删除")
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_MoveUp, "前移")
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_MoveDown, "后移")
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_SelMiniImage, "提取报告图")
        ElseIf intType = 1 Then
            Set cbrControl = .Add(xtpControlButton, comMenu_Cap_Process, "图像处理")
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_DevSet, "分页设置")
            Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportAll, "导出...")
            Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
        Else
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_DevSet, "分页设置")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Import, "导入...")
            Set cbrControl = .Add(xtpControlButton, conMenu_Manage_DeleteImage, "删除")
            Set cbrControl = .Add(xtpControlButton, conMenu_Manage_RefreshImg, "刷新")
        End If
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub

Private Sub ShowPopupMark()
    '------------------------------------------------
'功能：创建鼠标右键弹出菜单
'------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim cbrToolPopup As CommandBarPopup
    
    
    '鼠标右键弹出菜单
    Set cbrToolBar = cbrMain.Add("鼠标右键", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_DelMarks, "清除标注")
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub


Private Sub Form_Activate()
    '根据需要加载图像
    
    '注：在Form的Activate和Paint时间中必须调用LoadImages方法
    '因为如果只在Activate方法中调用LoadImages方法，可能造成报告图不会在第一时间显示，必须用鼠标点击一下报告图才会显示
    '如果只在Paint方法中调用LoadImages方法，由于该方法中使用了UnLoad卸载控件数组，可能造成“不能从该上下文中卸载”的错误
    
    Call LoadImages
End Sub

Private Sub Form_Load()
    
    '标记本次刷新已经加载图像
    blnLoadImages = True
    
    '标记窗体已经首次加载
    mblnIsInitFace = False
        
    mintMoustType = MarkType.自动编号

    
    '设置默认颜色
    lngColor(1) = RGB(186, 186, 186)
    lngColor(2) = RGB(255, 215, 0)
    lngColor(3) = RGB(255, 0, 255)
    lngColor(4) = RGB(255, 0, 130)
    lngColor(5) = RGB(0, 255, 0)
    lngColor(6) = RGB(130, 255, 255)
    lngColor(7) = RGB(255, 255, 0)
    lngColor(8) = RGB(0, 0, 255)
    lngColor(9) = RGB(0, 160, 0)
    
    '定义UIDRoot=1
    mdcmGlobal.RegString("UIDRoot") = "1"
    
End Sub

Public Sub MouseWheel(intDirection As Integer)
'处理鼠标滚轮的事件
'参数：intDirection --- 鼠标滚轮的方向；1--向上；0--向下
    
    On Error Resume Next
    
    If vscrollMini.Visible = False Then Exit Sub
    
    If intDirection = 1 Then '上翻一页
        If vscrollMini.value - 1 < 1 Then
            vscrollMini.value = 1
        Else
            vscrollMini.value = vscrollMini.value - 1
        End If
    Else        '下翻一页
        If vscrollMini.value + 1 > vscrollMini.Max Then
            vscrollMini.value = vscrollMini.Max
        Else
            vscrollMini.value = vscrollMini.value + 1
        End If
    End If
End Sub

Public Sub subDispScroll()
'------------------------------------------------
'功能：自动判断是否需要显示或隐藏滚动条
'返回：无，直接显示或隐藏滚动条。
'------------------------------------------------
    Dim ii As Integer
    
    If dcmMiniImageC.Images.Count > dcmMiniImageC.MultiColumns * dcmMiniImageC.MultiRows Then       '图像总数大于显示数，显示滚动条
        '摆放滚动条位置，并显示滚动条
        vscrollMini.Move dcmMiniImageC.Width - vscrollMini.Width, dcmMiniImageC.Top, vscrollMini.Width, dcmMiniImageC.Height
        vscrollMini.Visible = True
        vscrollMini.ZOrder
        vscrollMini.Refresh
        
        ''''''''''''''''''[关于滚动条需要单独仔细分析]'''''''''''''''''''''''''
        vscrollMini.Min = 1
        vscrollMini.Max = dcmMiniImageC.Images.Count - dcmMiniImageC.MultiColumns * dcmMiniImageC.MultiRows + 1
        If vscrollMini.Max < 1 Then vscrollMini.Max = 1
        vscrollMini.LargeChange = dcmMiniImageC.MultiColumns * dcmMiniImageC.MultiRows
        If dcmMiniImageC.CurrentIndex > vscrollMini.Max Then
            vscrollMini.value = vscrollMini.Max
            dcmMiniImageC.CurrentIndex = vscrollMini.Max
        Else
            vscrollMini.value = dcmMiniImageC.CurrentIndex
        End If
    Else                '图像数少于可显示数，隐藏滚动条
'        ii = dcmMiniature.Images.Count - dcmMiniature.MultiColumns * dcmMiniature.MultiRows + 1
'        If dcmMiniature.Images.Count - dcmMiniature.CurrentIndex + 1 < dcmMiniature.MultiColumns * dcmMiniature.MultiRows Then
'            dcmMiniature.CurrentIndex = IIf(ii < 1, 1, ii)
'        End If
'        vscrollMini.Value = dcmMiniature.CurrentIndex
        vscrollMini.Visible = False
    End If
    
    If vscrollMini.Visible = True Then
        dcmMiniImageC.Width = dcmMiniImageC.Width - vscrollMini.Width - 20
        
        vscrollMini.Height = dcmMiniImageC.Height - 40
        vscrollMini.Left = dcmMiniImageC.Width - 20
    Else
        dcmMiniImageC.Width = dcmMiniImageC.Width
    End If
End Sub


Private Sub InitLoaclParas()
    Dim strRegPath As String
    
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReportImage\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReportImage"
    End If
    
    ucMiniImageViewer.PageImgCount = Val(GetSetting("ZLSOFT", strRegPath, "报告缩略图数量", 5))
    
    '读取标记图区域，报告图区域 和缩略图区域的高度
    mlngCY1 = GetSetting("ZLSOFT", strRegPath, "CY1", 180)
    mlngMarkW = GetSetting("ZLSOFT", strRegPath, "MarkW", 300)
    mlngCY2 = GetSetting("ZLSOFT", strRegPath, "CY2", 400)
    mlngRptImgW = GetSetting("ZLSOFT", strRegPath, "RptImgW", 100)
    mlngCY3 = GetSetting("ZLSOFT", strRegPath, "CY3", 200)
End Sub

Private Sub Form_Paint()
    '根据需要加载图像
    Call LoadImages
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String
    
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReportImage\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReportImage"
    End If
    
    Call SaveSetting("ZLSOFT", strRegPath, "报告缩略图数量", ucMiniImageViewer.PageImgCount)
    
    '保存标记图区域，报告图区域和缩略图区域的高度
    '285是Pane的标题高度，使用了标题，就需要加回这个高度
    SaveSetting "ZLSOFT", strRegPath, "CY1", picMark.Height + 285
    SaveSetting "ZLSOFT", strRegPath, "MarkW", picMark.Width
    SaveSetting "ZLSOFT", strRegPath, "CY2", picReportImage.Height + 285
    SaveSetting "ZLSOFT", strRegPath, "RptImgW", picReportImage.Width
    SaveSetting "ZLSOFT", strRegPath, "CY3", picMiniImageC.Height + 285
    
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReport\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReport"
    End If
    SaveSetting "ZLSOFT", strRegPath, "CX3", Me.Width
    SaveSetting "ZLSOFT", strRegPath, "CY3", Me.Height
End Sub

Private Sub picMark_Resize()
    If picMark.Height = 0 Or picMark.Width = 0 Then Exit Sub
    
    On Error Resume Next
    
    '判断宽高比
    If picMark.Width / picMark.Height > 2 Then  '数字标记放在右边
        dcmMark.Left = 0
        dcmMark.Top = 0
        dcmMark.Width = Abs(picMark.ScaleWidth - picNumMark.ScaleWidth - 50)
        dcmMark.Height = picMark.ScaleHeight
        
        picNumMark.Left = dcmMark.Width
        If picMark.Height > picNumMark.Height Then
            picNumMark.Top = (picMark.ScaleHeight - picNumMark.ScaleHeight) / 2
        Else
            picNumMark.Top = 0
        End If
    Else    '数字标记放在下面
        dcmMark.Left = 0
        dcmMark.Top = 0
        dcmMark.Width = picMark.ScaleWidth
        dcmMark.Height = Abs(picMark.ScaleHeight - picNumMark.ScaleHeight - 50)
        
        If picMark.Width > picNumMark.Width Then
            picNumMark.Left = (picMark.ScaleWidth - picNumMark.ScaleWidth) / 2
        Else
            picNumMark.Left = 0
        End If
        picNumMark.Top = dcmMark.Height
    End If
End Sub

Private Sub picMiniCache_Resize()
On Error Resume Next
    cboCache.Left = 0
    cboCache.Top = 0
    cboCache.Width = picMiniCache.Width
    
    ucMiniCache.Left = 0
    ucMiniCache.Top = cboCache.Top + cboCache.Height
    ucMiniCache.Width = picMiniCache.ScaleWidth
    ucMiniCache.Height = picMiniCache.ScaleHeight - ucMiniCache.Top
End Sub

Private Sub picMiniImageC_Resize()
'    If picMiniImage.Width < 50 Or picMiniImage.Height < 50 Then Exit Sub
'    dcmMiniImage.Left = 0
'    dcmMiniImage.Top = 0
'    dcmMiniImage.Width = picMiniImage.Width - 50
'    dcmMiniImage.Height = picMiniImage.Height - 50

    Dim iRows As Integer
    Dim iCols As Integer
    
    On Error Resume Next
    
    dcmMiniImageC.Left = 0
    dcmMiniImageC.Top = 0
    dcmMiniImageC.Width = picMiniImageC.Width
    dcmMiniImageC.Height = picMiniImageC.Height
    
    '自动对图像做布局
    '计算缩略图的图像布局
    If mintShowPhotoNumber < dcmMiniImageC.Images.Count Then
        ResizeRegion mintShowPhotoNumber, dcmMiniImageC.Width, picMiniImageC.Height, iRows, iCols
    Else
        ResizeRegion dcmMiniImageC.Images.Count, dcmMiniImageC.Width, dcmMiniImageC.Height, iRows, iCols
    End If
    
    dcmMiniImageC.MultiColumns = iCols
    dcmMiniImageC.MultiRows = iRows
    '处理滚动条
    'If vscrollMini.Visible = True Then
    dcmMiniImageC.Width = picMiniImageC.Width - vscrollMini.Width - 20
    
    vscrollMini.Height = dcmMiniImageC.Height - 40
    vscrollMini.Left = dcmMiniImageC.Width - 20
    'End If
End Sub


Private Sub InitFaceScheme()
    '初始界面布局
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, pane4 As Pane
    
    With dkpMain
        .CloseAll
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    If mintShowMarkImage = 1 Then
        picMark.Visible = True
        dcmMark.Visible = True
        picNumMark.Visible = True
        
        Set Pane1 = dkpMain.CreatePane(1, mlngMarkW, mlngCY1, DockTopOf, Nothing)
        Pane1.Title = "标记图"
        Pane1.Handle = picMark.hWnd
        Pane1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        '根据宽高比，摆放报告图的位置
        If ((mlngCY1 = mlngCY2) And (mlngMarkW + mlngRptImgW > mlngCY1)) _
            Or (((mlngCY1 <> mlngCY2)) And (mlngMarkW + mlngRptImgW > mlngCY1 + mlngCY2)) Then
            Set Pane2 = dkpMain.CreatePane(2, mlngRptImgW, mlngCY2, DockLeftOf, Pane1)
        Else
            Set Pane2 = dkpMain.CreatePane(2, mlngRptImgW, mlngCY2, DockBottomOf, Pane1)
        End If
    Else
        picMark.Visible = False
        dcmMark.Visible = False
        picNumMark.Visible = False
        
        Set Pane2 = dkpMain.CreatePane(2, mlngRptImgW, mlngCY2, DockTopOf, Nothing)
    End If
    
'    If mobjStudyImage Is Nothing Then
'        Set mobjStudyImage = New clsStudyImages
'    End If
    
    mblnUseActiveVideo = False
    If mlngModule = G_LNG_VIDEOSTATION_MODULE Or mlngModule = G_LNG_PATHSTATION_MODULE Then
        mblnUseActiveVideo = GetSetting("ZLSOFT", "公共模块", "UseActiveVideo", "true")
        Call SaveSetting("ZLSOFT", "公共模块", "UseActiveVideo", mblnUseActiveVideo)
    End If

    Pane2.Title = "报告图"
    Pane2.Handle = picReportImage.hWnd
    Pane2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    picMiniImageC.Visible = Not mblnUseActiveVideo
    picMiniViewer.Visible = mblnUseActiveVideo

    Set Pane3 = dkpMain.CreatePane(3, 0, mlngCY3, DockBottomOf, Nothing)
    Pane3.Title = "缩略图"
    Pane3.Handle = IIf(mblnUseActiveVideo, picMiniViewer.hWnd, picMiniImageC.hWnd)
    Pane3.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    If mblnUseAfterCapture And mlngModule <> 1290 Then
        Set pane4 = dkpMain.CreatePane(4, 0, mlngCY3, DockBottomOf, Nothing)
        pane4.Title = "缓存图"
        pane4.Handle = picMiniCache.hWnd
        pane4.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        pane4.AttachTo Pane3
    Else
        picMiniCache.Visible = False
    End If
    
    mblnIsInitFace = True
End Sub

Private Function LoadMiniCache(Optional ByVal lngEventType As TVideoEventType = vetUpdateImg) As Boolean
    Dim i As Integer
    Dim strQueryValue As String
    Dim strSql As String
    
    strSql = "select A.姓名,A.检查号,A.性别,A.年龄,A.接收日期 As 检查日期,A.检查UID,B.序列UID,B.序列号 " & _
            "from 影像临时记录 A, 影像临时序列 B where A.检查uid = B.检查uid And A.接收日期 Between Sysdate-7 And Sysdate " & _
            "order by 接收日期 desc"
            
    Set mrsImageCache = zlDatabase.OpenSQLRecord(strSql, "")

    cboCache.Clear
    
    For i = 0 To mrsImageCache.RecordCount - 1
        If i = 0 Then strQueryValue = Nvl(mrsImageCache!序列UID)
        
        cboCache.AddItem "姓名：" & Nvl(mrsImageCache!姓名) & "  检查号：" & Nvl(mrsImageCache!检查号) & "  序列" & Nvl(mrsImageCache!序列号)
        If mstrTmpQueryValue = Nvl(mrsImageCache!序列UID) And lngEventType <> vetAfterUpdateImg Then cboCache.ListIndex = i
        
        mrsImageCache.MoveNext
    Next
    
    ucMiniCache.ImgViewer.Images.Clear
    ucMiniCache.ShowCheckBox = 1
    
    If mrsImageCache.RecordCount > 0 Or lngEventType = vetAfterUpdateImg Then
        If cboCache.ListIndex < 0 And mrsImageCache.RecordCount > 0 Then
            cboCache.ListIndex = 0
            mstrTmpQueryValue = strQueryValue
        Else
            cboCache_Click
        End If
    End If
End Function

Private Function LoadMiniImages() As Boolean
    Dim lngMsgHwnd As Long
    
    
    If mblnUseActiveVideo Then
'        lngMsgHwnd = mobjStudyImage.hWnd
'
'        Call mobjStudyImage.RefreshImages(mlngAdviceID, mlngAdviceID, mblnMoved, True)
        ucMiniImageViewer.ShowCheckBox = 1
        Call ucMiniImageViewer.RefreshImage(0, mlngAdviceID, mblnMoved, True)
    Else
        Call GetRptImages(dcmMiniImageC, mlngAdviceID, mblnMoved)
    
        Call AdjustDicomViewerLayout
    End If

End Function


Private Sub AdjustDicomViewerLayout()
'------------------------------------------------
'功能：将图像添加到缩略图dcmMiniature中
'参数：img－－输入的DICOM图像
'返回：无，直接将图像添加到缩略图dcmMiniature中
'------------------------------------------------
    Dim iRows As Integer
    Dim iCols As Integer
    
    '计算缩略图的图像布局
    If mintShowPhotoNumber < dcmMiniImageC.Images.Count + 1 Then
        ResizeRegion mintShowPhotoNumber, dcmMiniImageC.Width, dcmMiniImageC.Height, iRows, iCols
    Else
        ResizeRegion dcmMiniImageC.Images.Count + 1, dcmMiniImageC.Width, dcmMiniImageC.Height, iRows, iCols
    End If
            
    dcmMiniImageC.MultiColumns = iCols
    dcmMiniImageC.MultiRows = iRows

    
'    '根据缩略图的检查UID和序列UID，修改img的值
'    subUniteUID img
'    dcmMiniature.Images.Add img
    
'    '处理缩略图中被选中的状态
'    If mintCurImgIndex > 0 And mintCurImgIndex <= dcmMiniature.Images.Count Then
'        dcmMiniature.Images(mintCurImgIndex).BorderColour = vbWhite
'    End If
    
    If dcmMiniImageC.Images.Count > 0 Then
         With dcmMiniImageC.Images(1)
            .BorderWidth = 1
            .BorderStyle = 6
            .BorderColour = vbRed
        End With
    End If
    
    mintCurImgIndex = 1
    
'    mintCurImgIndex = dcmMiniature.Images.Count
    '显示滚动条
    Call subDispScroll
End Sub


Private Sub LoadReportImages()
    
    On Error GoTo errH
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim cTable As cEPRTable, oPicture As StdPicture
    Dim strPicFile As String
    Dim iRImageCount As Integer
    Dim objFile As New Scripting.FileSystemObject
    Dim blnGetTable As Boolean
    Dim i As Integer
    
    '初始化各个对象
    Call ClearReportImages
        
    '如果存在报告内容，则从报告内容中读取报告图和标记图，否则从报告单格式中读取标记图
    If mlngReportID <> 0 Then
        strSql = "Select Id As 表格Id From 电子病历内容" & vbNewLine & _
            " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
            " Order By 对象序号"
        If mblnMoved = True Then
            strSql = Replace(strSql, "电子病历内容", "H电子病历内容")
        End If

        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngReportID)
    Else
        strSql = "Select Id As 表格Id From 病历文件结构" & vbNewLine & _
            " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
            " Order By 对象序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngFileID)
    End If
    
    iRImageCount = 0
    Do While Not rsTemp.EOF
    
        Set cTable = New cEPRTable
        If mlngReportID = 0 Then
            blnGetTable = cTable.GetTableFromDB(cprET_病历文件定义, mlngFileID, Val("" & rsTemp!表格ID))
        Else
            blnGetTable = cTable.GetTableFromDB(cprET_单病历审核, mlngReportID, Val("" & rsTemp!表格ID))
        End If
        If blnGetTable Then
            
            '记录图像所在表格ID
            If pTableID = "" Then
                pTableID = cTable.ID
            Else
                pTableID = pTableID & ";" & cTable.ID
            End If
            
            '创建viewer
            iRImageCount = iRImageCount + 1
            Load dcmReportImage(iRImageCount)
            dcmReportImage(iRImageCount).BorderStyle = 1
            dcmReportImage(iRImageCount).Labels.AddNew
            dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).LabelType = doLabelRectangle
            dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).Left = 1
            dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).Top = 1
            dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).LineWidth = 2
            dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).ForeColour = vbWhite
            dcmReportImage(iRImageCount).Visible = True
            
            mSelViewerIndex = iRImageCount

            '记录图像框的宽度和高度，该宽高比例用于后续对图像行列布局
            If cTable.ExtendTag <> "" Then
                If Val(Split(cTable.ExtendTag, "|")(0)) = 0 Then
                    dcmReportImage(iRImageCount).tag = cTable.Width & "|" & cTable.Height
                Else
                    dcmReportImage(iRImageCount).tag = (cTable.Width - Val(Split(cTable.ExtendTag, "|")(1)) - 30) & "|" & cTable.Height
                End If
            Else
                dcmReportImage(iRImageCount).tag = cTable.Width & "|" & cTable.Height
            End If
            
            
            For i = 1 To cTable.Pictures.Count
                strPicFile = App.Path & "\PACSPic" & i & ".JPG"
                If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True

                Set oPicture = cTable.Pictures(i).OrigPic
                SavePicture oPicture, strPicFile
                If objFile.FileExists(strPicFile) Then
                    '显示标记图和报告图
                    If cTable.Pictures(i).PictureType = EPRMarkedPicture And dcmMark.Images.Count = 0 Then

                        '只处理第一个标记图
                        dcmMark.Images.AddNew
                        
                        dcmMark.Images(1).FileImport strPicFile, "BMP"
                        dcmMark.Images(1).tag = cTable.Pictures(i).ID
                        '保存标记图基础数据
                        Set pobjMarks = cTable.Pictures(i).PicMarks
                        pMarkImageID = cTable.Pictures(i).ID

                        mdblMarkZoom = dcmMark.Images(1).SizeX / cTable.Pictures(i).Width * Screen.TwipsPerPixelX
                        '显示标注
                        If cTable.Pictures(i).PicMarks.Count > 0 Then
                            drawPicMarks dcmMark.Images(1), cTable.Pictures(i).PicMarks
                        End If
                    Else

                        dcmReportImage(iRImageCount).Images.AddNew
                        dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).FileImport strPicFile, "BMP"
                        If cTable.Pictures(i).PicName = "" Then
                            dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).tag = mdcmGlobal.NewUID & ".jpg"
                        Else
                            dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).tag = cTable.Pictures(i).PicName
                        End If
                        
                        dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).BorderWidth = 3
                        dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).BorderColour = vbWhite
                        dcmReportImage(iRImageCount).CurrentIndex = 1
                        mselReportImgIndex = 1
                    End If
                    '删除临时图像
                    Kill strPicFile
                End If
            Next
        End If
        
        rsTemp.MoveNext
    Loop
    If dcmReportImage.Count > 1 Then dcmReportImage(1).Labels(1).ForeColour = vbRed
    Call picReportImage_Resize
    

Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function DecideMarkImagesVisible() As Integer
'------------------------------------------------
'功能：判断当前选中检查标记图是否可见
'参数：无
'返回：int类型，1-显示标记图  2-隐藏标记图
'-----------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim cTable As cEPRTable, oPicture As StdPicture
    Dim strPicFile As String
    Dim objFile As New Scripting.FileSystemObject
    Dim blnGetTable As Boolean
    Dim i As Integer
    
        
    '如果存在报告内容，则从报告内容中读取数据，否则从报告单格式中读取数据
    If mlngReportID <> 0 Then
        strSql = "Select Id As 表格Id From 电子病历内容" & vbNewLine & _
            " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
            " Order By 对象序号"
        If mblnMoved = True Then
            strSql = Replace(strSql, "电子病历内容", "H电子病历内容")
        End If

        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "报告内容中读取", mlngReportID)
    Else
        strSql = "Select Id As 表格Id From 病历文件结构" & vbNewLine & _
            " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
            " Order By 对象序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "从报告单格式中读取", mlngFileID)
    End If
    
    Do While Not rsTemp.EOF
        Set cTable = New cEPRTable
        If mlngReportID = 0 Then
            blnGetTable = cTable.GetTableFromDB(cprET_病历文件定义, mlngFileID, Val("" & rsTemp!表格ID))
        Else
            blnGetTable = cTable.GetTableFromDB(cprET_单病历审核, mlngReportID, Val("" & rsTemp!表格ID))
        End If
        
        If blnGetTable Then
            If cTable.Pictures.Count > 0 Then
                For i = 1 To cTable.Pictures.Count
                    If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                        DecideMarkImagesVisible = 1
                        Exit Do
                    Else
                        DecideMarkImagesVisible = 0
                    End If
                Next
            Else
                DecideMarkImagesVisible = 0
            End If
        End If
        rsTemp.MoveNext
    Loop

End Function


Private Sub drawPicMarks(img As DicomImage, thisMarks As cPicMarks)
'显示标注，只支持数字编号标注
    Dim i As Integer
    Dim iLabelCount As Integer
    
    img.Labels.Clear
    For i = 1 To thisMarks.Count
        If thisMarks(i).类型 = 6 Then   '圆形编号
            With thisMarks(i)
                img.Labels.AddNew
                iLabelCount = img.Labels.Count
                img.Labels(iLabelCount).LabelType = doLabelEllipse
                img.Labels(iLabelCount).BackColour = IIf(.填充色 = 0, vbYellow, .填充色)
                img.Labels(iLabelCount).Transparent = False
                img.Labels(iLabelCount).Left = .X1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Top = .Y1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Width = 20
                img.Labels(iLabelCount).Height = 20
                img.Labels(iLabelCount).ImageTied = True
                
                img.Labels.AddNew
                iLabelCount = img.Labels.Count
                img.Labels(iLabelCount).LabelType = doLabelEllipse
                img.Labels(iLabelCount).ForeColour = vbBlack
                img.Labels(iLabelCount).Transparent = True
                img.Labels(iLabelCount).Left = .X1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Top = .Y1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Width = 20
                img.Labels(iLabelCount).Height = 20
                img.Labels(iLabelCount).ImageTied = True

                img.Labels.AddNew
                iLabelCount = img.Labels.Count
                img.Labels(iLabelCount).LabelType = doLabelText
                img.Labels(iLabelCount).Transparent = True
                img.Labels(iLabelCount).ForeColour = vbBlack
                img.Labels(iLabelCount).FontSize = 11
                img.Labels(iLabelCount).FontName = "Arial Bold"
                img.Labels(iLabelCount).Left = .X1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Top = .Y1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).AutoSize = True
                img.Labels(iLabelCount).Text = .内容
                img.Labels(iLabelCount).ImageTied = True
            End With
        End If
    Next i
End Sub
 
Private Sub picMiniViewer_Resize()
On Error Resume Next
    ucMiniImageViewer.Left = 0
    ucMiniImageViewer.Top = 0
    ucMiniImageViewer.Width = picMiniViewer.ScaleWidth
    ucMiniImageViewer.Height = picMiniViewer.ScaleHeight
End Sub

Private Sub picReportImage_Resize()
    Dim i As Integer
    Dim rectH As Long, rectW As Long    '图象框可以使用的区域宽高
    Dim picH As Long, picW As Long      '图像实际宽高，作为比例使用
    Dim iCols As Integer, iRows As Integer
    Dim dImg As DicomImage
    
    If dcmReportImage.Count = 1 Then Exit Sub
    
    On Error Resume Next
    
    '首先计算每个图象框可占用的最大宽高
    
    rectH = picReportImage.Height / (dcmReportImage.Count - 1)
    rectW = picReportImage.Width
    If rectH < 100 Or rectW < 100 Then Exit Sub
    
    For i = 1 To dcmReportImage.Count - 1
        '按照图像比例，计算图象框的真实宽度和高度
        picW = Val(Split(dcmReportImage(i).tag, "|")(0))
        picH = Val(Split(dcmReportImage(i).tag, "|")(1))
        
        dcmReportImage(i).Height = rectH - 100
        dcmReportImage(i).Width = rectW - 100
        
        dcmReportImage(i).Left = 0
        dcmReportImage(i).Top = rectH * (i - 1)
        
        dcmReportImage(i).Labels(1).Width = Abs(dcmReportImage(i).Width / Screen.TwipsPerPixelX - 2)
        dcmReportImage(i).Labels(1).Height = Abs(dcmReportImage(i).Height / Screen.TwipsPerPixelY - 1)

        
        '调整图像显示布局
        ResizeRegion dcmReportImage(i).Images.Count, picW, picH, iRows, iCols
        dcmReportImage(i).MultiColumns = iCols
        dcmReportImage(i).MultiRows = iRows
    Next i
End Sub

Public Sub zlChangeFormat(FormatID As Long)

    On Error GoTo errH
    
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim cTable As cEPRTable, oPicture As StdPicture
    Dim strPicFile As String
    Dim iRImageCount As Integer
    Dim blnHasMarkImage As Boolean
    Dim blnGetTable As Boolean
    Dim objFile As New Scripting.FileSystemObject
    Dim i As Integer
    
    '提示格式变换中图象框、标记图等的数量变化
    If FormatID = 0 Then     '标准格式，查 病历文件结构
        strSql = "Select Id As 表格Id From 病历文件结构" & vbNewLine & _
        "       Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By 对象序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngFileID)
    Else        '范文格式，查 病历范文内容
        strSql = "Select Id As 表格Id From 病历范文内容" & vbNewLine & _
        "       Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By 对象序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, FormatID)
    End If
    If rsTemp.RecordCount > 0 Then
        If rsTemp.RecordCount < dcmReportImage.Count - 1 Then
            If MsgBoxD(Me, "新格式中图象框数量少于当前格式，当前的部分图象框会被删除，是否更换格式？", vbOKCancel) = vbCancel Then
                Exit Sub
            Else
                '先删除多余的图象框
                For i = dcmReportImage.Count - 1 - rsTemp.RecordCount To 1 Step -1
                    Unload dcmReportImage(dcmReportImage.Count - 1)
                Next i
            End If
        End If
        
        '读取图象框中的标记图和报告图
        iRImageCount = 0
        pTableID = ""
        Do While Not rsTemp.EOF
            Set cTable = New cEPRTable
            If FormatID = 0 Then
                blnGetTable = cTable.GetTableFromDB(cprET_病历文件定义, mlngFileID, Val("" & rsTemp!表格ID))
            Else
                blnGetTable = cTable.GetTableFromDB(cprET_全文示范编辑, FormatID, Val("" & rsTemp!表格ID))
            End If
            If blnGetTable Then
                iRImageCount = iRImageCount + 1
                If iRImageCount > dcmReportImage.Count - 1 Then
                    '创建Viewer
                    Load dcmReportImage(iRImageCount)
                    dcmReportImage(iRImageCount).BorderStyle = 1
                    dcmReportImage(iRImageCount).Labels.AddNew
                    dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).LabelType = doLabelRectangle
                    dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).Left = 1
                    dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).Top = 1
                    dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).LineWidth = 2
                    dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).ForeColour = vbWhite
                    dcmReportImage(iRImageCount).Visible = True
                End If
                mSelViewerIndex = iRImageCount
                
                '记录图像框的宽度和高度，该宽高比例用于后续对图像行列布局
                If cTable.ExtendTag <> "" Then
                    If Val(Split(cTable.ExtendTag, "|")(0)) = 0 Then
                        dcmReportImage(iRImageCount).tag = cTable.Width & "|" & cTable.Height
                    Else
                        dcmReportImage(iRImageCount).tag = (cTable.Width - Val(Split(cTable.ExtendTag, "|")(1)) - 30) & "|" & cTable.Height
                    End If
                Else
                    dcmReportImage(iRImageCount).tag = cTable.Width & "|" & cTable.Height
                End If
                
                '更新标记图
                For i = 1 To cTable.Pictures.Count
                    strPicFile = App.Path & "\PACSPic" & i & ".JPG"
                    If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True

                    Set oPicture = cTable.Pictures(i).OrigPic
                    SavePicture oPicture, strPicFile
                    If objFile.FileExists(strPicFile) Then
                        '显示标记图
                        If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                            blnHasMarkImage = True
                            '先清除当前标记图，再更新
                            dcmMark.Images.Clear
                            dcmMark.Images.AddNew
                            dcmMark.Images(1).FileImport strPicFile, "BMP"
                            dcmMark.Images(1).tag = cTable.Pictures(i).ID
                            '如果当前没有标记，则读取新格式中标记图的标记
                            If pobjMarks Is Nothing Then
                                Set pobjMarks = cTable.Pictures(i).PicMarks
                            End If
                            pMarkImageID = cTable.Pictures(i).ID
    
                            mdblMarkZoom = dcmMark.Images(1).SizeX / cTable.Pictures(i).Width * Screen.TwipsPerPixelX
                            '显示标注
                            If pobjMarks.Count > 0 Then
                                drawPicMarks dcmMark.Images(1), pobjMarks
                            End If
                        End If
                        '删除临时图像
                        Kill strPicFile
                    End If
                Next
            End If
            rsTemp.MoveNext
        Loop
    End If
    
    If blnHasMarkImage = False Then
        '当前格式没有标记图，删除当前显示的标记图
        pMarkImageID = 0
        
        dcmMark.Images.Clear
        If Not (pobjMarks Is Nothing) Then
            For i = 1 To pobjMarks.Count
                pobjMarks.Remove 1
            Next i
        End If
    End If
    Call picReportImage_Resize

Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ucMiniCache_OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucMiniCache.ImgViewer.Images.Count <= 0 Then Exit Sub
    If Button = 1 Then ucMiniCache.ImgChecked(ucMiniCache.SelectIndex) = Not ucMiniCache.ImgChecked(ucMiniCache.SelectIndex)
End Sub

Private Sub ucMiniCache_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucMiniCache.ImgViewer.Images.Count <= 0 Then Exit Sub
    If Button = 2 Then Call ShowPopupImage(2)
End Sub

Private Sub ucMiniImageViewer_OnDbClick(ByVal lngSelectedIndex As Long, blnContinue As Boolean)
    If ucMiniImageViewer.CurImageCount > 0 And mSelViewerIndex <> 0 And dcmReportImage.Count > 1 Then
        '判断当前双击的操作动作
        If mintImageDblClick = 0 Then   '直接写入报告
            Dim dcmImage As DicomImage
            Set dcmImage = mSelMiniImg
            
            '调用将当前图形添加到图象框过程
            Call DcmAddImage(dcmImage, mSelViewerIndex)
            
        Else                            '先打开图片编辑窗口
            '先关闭大图窗口
            ReleaseCapture      '解锁鼠标
            frmShowImg.HideMe
            
            Call OpenImageProcessWind
        End If

    End If
    
    blnContinue = False
End Sub

Private Sub ucMiniImageViewer_OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucMiniImageViewer.ImgViewer.Images.Count <= 0 Then Exit Sub
    If Button = 1 Then ucMiniImageViewer.ImgChecked(ucMiniImageViewer.SelectIndex) = Not ucMiniImageViewer.ImgChecked(ucMiniImageViewer.SelectIndex)
End Sub

Private Sub ucMiniImageViewer_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucMiniImageViewer.ImgViewer.Images.Count <= 0 Then Exit Sub
    If Button = 2 Then Call ShowPopupImage(1)
End Sub

Private Sub ucMiniImageViewer_OnSelChange(ByVal lngSelectedIndex As Long)
    Set mSelMiniImg = ucMiniImageViewer.SelectImage
End Sub

Private Sub vscrollMini_Change()
    Dim iImgIndex As Integer
    
    If dcmMiniImageC.Images.Count > 0 And (mintCurImgIndex > 0 And mintCurImgIndex <= dcmMiniImageC.Images.Count) Then
        iImgIndex = vscrollMini.value + mintCurImgIndex - dcmMiniImageC.CurrentIndex
        If iImgIndex <= 0 Then
            iImgIndex = 1
        ElseIf iImgIndex > dcmMiniImageC.Images.Count Then
            iImgIndex = dcmMiniImageC.Images.Count
        End If
        dcmMiniImageC.CurrentIndex = vscrollMini.value
        
        dcmMiniImageC.Images(mintCurImgIndex).BorderColour = vbWhite
        mintCurImgIndex = iImgIndex
        dcmMiniImageC.Images(mintCurImgIndex).BorderColour = vbRed
    End If

End Sub

Private Sub LoadImages()
'------------------------------------------------
'功能：加载报告图和缩略图
'参数：
'返回：无，直接加载图像，并修噶 blnLoadImages状态
'-----------------------------------------------
    '如果本次刷新没有加载图像，则加载图像
    If blnLoadImages = False Then
        '读取后台采集的图像
        If mblnUseAfterCapture And mlngModule <> 1290 Then
            Call LoadMiniCache
        End If
        
        '读取和显示当前可选报告图像
        Call LoadMiniImages
        '根据报告单格式，或者报告内容格式，读取标记图和报告图
        Call LoadReportImages
        '标记本次刷新已经加载图像
        blnLoadImages = True
    End If
End Sub
