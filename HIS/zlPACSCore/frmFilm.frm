VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmFilm 
   Caption         =   "胶片打印预览"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11400
   Icon            =   "frmFilm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimePass 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1560
      Top             =   120
   End
   Begin VB.PictureBox picBak 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5568
      Left            =   720
      ScaleHeight     =   5565
      ScaleWidth      =   9135
      TabIndex        =   0
      Top             =   720
      Width           =   9135
      Begin VB.PictureBox picFilm 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H00000000&
         Height          =   4056
         Left            =   1800
         ScaleHeight     =   4020
         ScaleWidth      =   5340
         TabIndex        =   2
         Top             =   570
         Width           =   5376
         Begin DicomObjects.DicomViewer FilmViewer 
            DragIcon        =   "frmFilm.frx":000C
            Height          =   1140
            Index           =   0
            Left            =   3720
            TabIndex        =   3
            Top             =   2520
            Visible         =   0   'False
            Width           =   1230
            _Version        =   262147
            _ExtentX        =   2159
            _ExtentY        =   2011
            _StockProps     =   35
            BackColor       =   0
            UseScrollBars   =   0   'False
         End
      End
      Begin VB.VScrollBar VScro 
         Height          =   3888
         Left            =   4
         Max             =   1
         TabIndex        =   1
         Top             =   672
         Width           =   250
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   7875
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2302
            MinWidth        =   2293
            Text            =   "中联软件"
            TextSave        =   "中联软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12647
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "大写"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   926
            MinWidth        =   706
            Text            =   "数字"
            TextSave        =   "NUM"
            Object.ToolTipText     =   "数字"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListMouse 
      Left            =   240
      Top             =   120
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":0CD6
            Key             =   "调窗"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":19B0
            Key             =   "框选缩放"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":1CCA
            Key             =   "前"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":1FE4
            Key             =   "漫游"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":2CBE
            Key             =   "后"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":2FD8
            Key             =   "缩放"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":38B2
            Key             =   "裁剪"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":3BCC
            Key             =   "左"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":3EE6
            Key             =   "右"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":4200
            Key             =   "上"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":451A
            Key             =   "下"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager ImgIcons 
      Left            =   240
      Top             =   2640
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmFilm.frx":4834
   End
   Begin XtremeCommandBars.CommandBars CommBar_Film 
      Left            =   300
      Top             =   1080
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmFilm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'-----------------------------------------------------
'窗体事件
'-----------------------------------------------------
Public Event AfterPrinted(strImageUIDS As String)      '打印完成

'''''''带事件的子窗体'''''''''''''''
Public WithEvents mfrmFilmView As frmFilmView
Attribute mfrmFilmView.VB_VarHelpID = -1

'窗体的公共变量-------------------------------------

Public clsTruePrinter As clsDicomPrint      ''DICOM打印机的设置
Public SelectedImage As DicomImage          ''记录当前被选中的图像，提供给模块统一做窗宽窗位功能按钮
Public intMouseState As Integer             ''记录鼠标的状态：0－无；1－调窗；2－漫游；3－缩放;4-无;5-框选缩放;6-裁剪:7-文字标注，跟frmFilmView窗体有交互操作
Public pstrSideMarker As String             ''记录当前需要标注的体位文字，跟frmFilmView窗体有交互操作
Public blnDefaultWW2 As Boolean             ''记录双窗口状态，提供给模块统一做窗宽窗位功能按钮
Public f As frmViewer

'窗体的私有变量'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private imgsPrint As New DicomImages        ''参与打印的图像正本集合
Private mblnPrinted As Boolean              ''记录是否有图像已经被打印过
Private mintPrintFilmCount As Integer       ''记录打印成功的次数
Private mintCellSpacing As Integer          ''显示胶片预览的时候，图片之间的间距，显示用，跟实际打印无关
Private mintFilmHeight As Integer           ''胶片的高度，单位英寸
Private mintFilmWidth As Integer            ''胶片的宽度，单位英寸
Private mblnIsPortrait As Boolean           ''是否纵向打印
Private mblnIsRow As Boolean                ''当前页面中，是否列优先
Private mblnIsCustom As Boolean             ''当前页面中，是否有行列自定义
Private mdubFilmRate As Double              ''胶片的长宽比例
Private mdubScreenRate As Double            ''屏幕的长宽比例
Private mblnBegin As Boolean

Private marrRCCount() As Integer            ''当前页面中，每行/每列的图像数目，在STANDARD情况下，aRCCount(1)表示列数
Private marrPages() As FilmType             ''每一页中的Viewer数量和DICOM标准布局方式，数组维度是胶片页数

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private mintSelectedViewer As Integer       ''记录当前被选中的Viewer序号
Private mintSelectedImage As Integer        ''记录当前被选中的图像序号
Private mintBaseX As Long                   ''记录鼠标原来的X位置
Private mintBaseY As Long                   ''记录鼠标原来的Y位置
Private mdcmSelectLabel As DicomLabel       ''当前被选中的标注
Private mblnDcmViewDown As Boolean          ''用于判断dcmView中鼠标是否被按下
Private mblnLabelMoving As Boolean          ''正在移动裁剪框
Private mblnCheckPrinter As Boolean         ''是否检查打印机的状态
Private mblnInTest As Boolean               ''记录是否处于测试状态
Private mintPageRange As Integer            ''打印页数范围：0-全部，1-当前页
Private mintTBMainPosition As Integer       ''记录主工具栏位置
Private mintTBImageProcessPosition As Integer   ''记录图像操作工具栏位置
Private mblnPrinting As Boolean             ''是否正在打印的过程中
Private mblnClearAfterPrint As Boolean      ''是否打印后清空图像

''''''''''''''''裁剪''''''''''''''''''''''''''''''''''''''''''''''
Private mintCutOutViewer                    '裁剪框所在的viewer序号
Private mintCutOutImage                     '裁剪框所在的图像序号
Private mintCutOutLabel                     '裁剪框所在的标注序号
Private mdblCutOutRatio As Double           '裁剪的比例，如果是固定比例则直接记录比例，无固定比例则为0
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'记录每个Viwer中图像的行数和列数
Private Type TLayout
    intRows As Integer
    intColumns As Integer
End Type

Private Type FilmType
    intViewerCount As Integer               '根据布局计算出来的Viewer总数量
    strPageFormat As String                 '这一页的布局，DICOM标准定义
    ViewerLayout() As TLayout               '这一页中，每个Viewer的行列布局
    intImageCount As Integer                '这一页中的图像总数
End Type

Private Type ImageSize
    intWidth As Integer                     '图片的最大宽度
    intHeight As Integer                    '图片的最大高度
End Type

'-----胶片打印用的TAG常量--------------------------------------------------
Private Const zlSpliter = "-ZL-"            'TAG中记录内容的分隔符
Private Const TAG_选择 = "1"
Private Const TAG_定位 = "2"
Private Const TAG_V宽度 = "3"
Private Const TAG_V高度 = "4"


'窗体的函数'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub subConfig()
'------------------------------------------------
'功能：将frmFilmConf窗体中设置的排版格式应用到胶片打印中。
'参数：无
'返回：无
'------------------------------------------------
    Dim i As Integer
    Dim strFilmFormat As String
    
    On Error GoTo err
    
    With frmFilmConf
        
        '记录胶片的宽度和高度
        i = InStr(.cobSize.Text, "X")
        mintFilmWidth = Val(Mid(.cobSize.Text, 1, i - 1))
        mintFilmHeight = Val(Mid(.cobSize.Text, i + 1))
        
        '记录胶片方向
        mblnIsPortrait = IIf(.cobAspect.Text = "纵向", True, False)       ''是否纵向打印
        
        'Option(0)--标准行列；Option(1)---行自定义；Option(2) ---列自定义
        
        '组合并记录胶片格式
        If Not .Option(2) Then        '行自定义或者标准格式
            If Not .Option(0) Then
                strFilmFormat = "ROW\" & .txtC(1)
            Else
                strFilmFormat = "STANDARD\" & .txtCol & "," & .txtRow
            End If
        Else            '列自定义
            strFilmFormat = "COL\" & .txtC(1)
        End If
        If Not .Option(0) Then
            If .Option(1) Then
                For i = 2 To Val(.txtRow)
                    strFilmFormat = strFilmFormat & "," & .txtC(i)
                Next i
            Else
                For i = 2 To Val(.txtCol)
                    strFilmFormat = strFilmFormat & "," & .txtC(i)
                Next i
            End If
        End If
    End With
    
    '设置菜单上对应的胶片规格和胶片格式
    Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, True).Text = frmFilmConf.cobSize.Text
    Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text = strFilmFormat
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub CommBar_Film_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer
    Dim intNowIndex As Integer
    Dim intTemp As Integer
    Dim thisControl As CommandBarControl
    Dim strLayout As String
    
    On Error GoTo err
    
    '''''''''''''''''''''''''''''[功能键设置窗宽窗位处理]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If control.Id >= 349 And control.Id <= 360 Then
        For i = 349 To 360
            If Not CommBar_Film.Item(3).FindControl(, i, , True) Is Nothing Then
                CommBar_Film.Item(3).FindControl(, i, , True).Checked = False
                If i = ID_Active_AdjustWindow_HandAdjustWindow_Custom Then
                    CommBar_Film.Item(3).FindControl(, i, , True).Checked = False
                End If
            End If
        Next
        control.Checked = True
        subFunctionWL CommBar_Film.Item(3).FindControl(, control.Id, , True), Me
        If control.Id = ID_Active_AdjustWindow_HandAdjustWindow_Custom Then
            CommBar_Film.Item(3).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow_Custom, , True).Checked = True
            CommBar_Film.Item(3).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow_Custom, , True).Checked = True
        End If
        
        Call subFilmViewButtonClick(CommBar_Film.Item(3).FindControl(, control.Id, , True))
        Exit Sub
    End If
    
    Select Case control.Id
    Case ID_frmFilm_FilmCol             '纵向
        If Not CommBar_Film.FindControl(, ID_frmFilm_FilmCol, , True).Checked Then
            CommBar_Film.FindControl(, ID_frmFilm_FilmCol, , True).Checked = True
            CommBar_Film.FindControl(, ID_frmFilm_FilmRow, , True).Checked = False
            mblnIsPortrait = True
            Call picBak_Resize
        End If
    Case ID_frmFilm_FilmRow             '横向
        If Not CommBar_Film.FindControl(, ID_frmFilm_FilmRow, , True).Checked Then
            CommBar_Film.FindControl(, ID_frmFilm_FilmRow, , True).Checked = True
            CommBar_Film.FindControl(, ID_frmFilm_FilmCol, , True).Checked = False
            mblnIsPortrait = False
            Call picBak_Resize
        End If
    Case ID_frmFilm_RectPhotCase        '正方形图像格
        CommBar_Film.FindControl(, ID_frmFilm_RectPhotCase, , True).Checked = Not control.Checked
        Call picBak_Resize
    Case ID_frmFilm_FormatCustom        '格式定义
        Call subSetFilmFormat
    Case ID_frmFilm_TakePictures        '照相
        Call CommBar_Execute_PrintFilm
        
    Case ID_frmFilm_FilmSize            '胶片大小
        i = InStr(control.Text, "X")
        If i <> 0 Then
            mintFilmWidth = Val(Mid(control.Text, 1, i - 1))
            mintFilmHeight = Val(Mid(control.Text, i + 1))
            Call picBak_Resize
        End If
    Case ID_frmFilm_Format              '胶片格式
        '修改胶片格式
        Call funChangeFormat(Me.VScro.Value, Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text)
        '重新显示这一页
        Call subShowOnePage(Me.VScro.Value)
    Case ID_frmFilm_Camera              '打印机
        Dim strThisFilmSize As String
        If Me.CommBar_Film.FindControl(, ID_frmFilm_Camera, , True).Text <> "" Then
            strThisFilmSize = cDICOMPrinter(Me.CommBar_Film.FindControl(, ID_frmFilm_Camera, , True).Text).strFilmSize
            Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, , True).Text = strThisFilmSize
            CommBar_Film_Execute Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, , True)
        End If
    Case ID_frmFilm_Quit                '退出
        Unload Me
    Case ID_frmFilm_OpenImages          ''打开图像
        Dim strImageIDs As String
        strImageIDs = frmPACSImg.zlOpenImages(Me, f)
        '打开图象
        Call OpenImages(strImageIDs)
    Case ID_frmFilm_DeleteImg            ''删除图像
        Call subDelImage
    Case ID_Active_AdjustWindow_HandAdjustWindow             ''调窗
        subCheckToolBar control
        intMouseState = IIf(control.Checked = True, 1, 0)
    Case ID_frmFilm_Pan                  ''漫游
        subCheckToolBar control
        intMouseState = IIf(control.Checked = True, 2, 0)
    Case ID_frmFilm_Zoom                 ''缩放
        subCheckToolBar control
        intMouseState = IIf(control.Checked = True, 3, 0)
    Case ID_frmFilm_FilterLengthUp       ''平滑增加
        If Not SelectedImage Is Nothing Then
            Call SubImageFiltering("miFilterLengthUp", SelectedImage)
            Call subFilmViewButtonClick(control)
        End If
    Case ID_frmFilm_FilterLengthDown     ''平滑减少
        If Not SelectedImage Is Nothing Then
            Call SubImageFiltering("miFilterLengthDown", SelectedImage)
            Call subFilmViewButtonClick(control)
        End If
    Case ID_frmFilm_RectZoom             ''框选缩放
        subCheckToolBar control
        intMouseState = IIf(control.Checked = True, 5, 0)
    Case ID_frmFilm_CutOut               ''裁剪
        subCheckToolBar control
        intMouseState = IIf(control.Checked = True, 6, 0)
        If intMouseState = 6 Then
            Call subCutOutClick
        End If
    Case ID_frmFilm_CutOut_Custom           ''自由比例裁剪
        Set thisControl = CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut, , True)
        Call subCheckToolBar(thisControl)
        thisControl.Checked = True
        Call subCutOutButtonState(control.Id)
        intMouseState = 6
        mdblCutOutRatio = 0
    Case ID_frmFilm_CutOut_14X17, ID_frmFilm_CutOut_11X14, ID_frmFilm_CutOut_10X14, _
        ID_frmFilm_CutOut_8X10, ID_frmFilm_CutOut_14X14, ID_frmFilm_CutOut_17X14, ID_frmFilm_CutOut_14X11, _
        ID_frmFilm_CutOut_14X10, ID_frmFilm_CutOut_10X8
        
        '固定比例裁剪,只要单击，就进入裁剪状态
        Set thisControl = CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut, , True)
        Call subCheckToolBar(thisControl)
        thisControl.Checked = True
        intMouseState = 6
        Call subCutOutButtonState(control.Id)
        Call subCutOutRatio(control.Id)
    Case ID_frmFilm_Invert               ''反白
        If Not SelectedImage Is Nothing Then
            Call subFlipRotate(SelectedImage, "Invert")
            Call subSynchronalImg(False, IMG_SYN_WINDOW)
            Call subFilmViewButtonClick(control)
        End If
    Case ID_frmFilm_RotateLeft           ''向左旋转90度
        If Not SelectedImage Is Nothing Then
            Call subFlipRotate(SelectedImage, "RotateAnticlockwise")
            Call subSynchronalImg(False, IMG_SYN_ROTATE)
            Call subFilmViewButtonClick(control)
        End If
    Case ID_frmFilm_RotateRight          ''向右旋转90度
        If Not SelectedImage Is Nothing Then
            Call subFlipRotate(SelectedImage, "RotateClockwise")
            Call subSynchronalImg(False, IMG_SYN_ROTATE)
            Call subFilmViewButtonClick(control)
        End If
    Case ID_frmFilm_FlipHorizontal       ''左右镜象
        If Not SelectedImage Is Nothing Then
            Call subFlipRotate(SelectedImage, "FlipHorizontal")
            Call subSynchronalImg(False, IMG_SYN_FLIP)
            Call subFilmViewButtonClick(control)
        End If
    Case ID_frmFilm_FlipVertical         ''上下镜象
        If Not SelectedImage Is Nothing Then
            Call subFlipRotate(SelectedImage, "FlipVertical")
            Call subSynchronalImg(False, IMG_SYN_FLIP)
            Call subFilmViewButtonClick(control)
        End If
    Case ID_frmFilm_Label_L             ''标注文字
        intMouseState = 7
        pstrSideMarker = "左"
        subCheckToolBar control
    Case ID_frmFilm_Label_R             ''标注文字
        intMouseState = 7
        pstrSideMarker = "右"
        subCheckToolBar control
    Case ID_frmFilm_Label_A             ''标注文字
        intMouseState = 7
        pstrSideMarker = "前"
        subCheckToolBar control
    Case ID_frmFilm_Label_P             ''标注文字
        intMouseState = 7
        pstrSideMarker = "后"
        subCheckToolBar control
    Case ID_frmFilm_Label_S             ''标注文字
        intMouseState = 7
        pstrSideMarker = "上"
        subCheckToolBar control
    Case ID_frmFilm_Label_I             ''标注文字
        intMouseState = 7
        pstrSideMarker = "下"
        subCheckToolBar control
    Case ID_frmFilm_Label_Delete        ''清除标注文字
        If Not SelectedImage Is Nothing Then
            For i = SelectedImage.Labels.Count To G_INT_SYS_LABEL_COUNT + 1 Step -1
                If SelectedImage.Labels(i).Text = "上" Or SelectedImage.Labels(i).Text = "下" Or _
                    SelectedImage.Labels(i).Text = "左" Or SelectedImage.Labels(i).Text = "右" Or _
                    SelectedImage.Labels(i).Text = "前" Or SelectedImage.Labels(i).Text = "后" Then
                    
                    SelectedImage.Labels.Remove (i)
                End If
            Next i
            SelectedImage.Refresh False
        End If
        '把标注过的图像上传到原始图集合中
        Call subReloadImgsPrint
    Case ID_frmFilm_Resume               ''恢复
        If Not SelectedImage Is Nothing Then
            Call subSynchronalImg(True, IMG_SYN_All)
            Call subFilmViewButtonClick(control)
        End If
    Case ID_frmFilm_SelAll              ''图像全选,自动图像同步
        Call SelAllImage(True)
    Case ID_frmFilm_SelSeries           ''选择当前序列
        Call SelOneSeries
    Case ID_frmFilm_SelInverse          ''反选
        Call SelectInverse
    Case ID_frmFilm_SelNone             ''全清
        Call SelAllImage(False)
    Case ID_frmFilm_Divide              ''图像分格
            '显示分格选择窗口
            strLayout = frmFilmLayout.ShowMe(Me)
            If Len(strLayout) = 3 And Val(left(strLayout, 1)) <> 0 And Val(Right(strLayout, 1)) <> 0 _
                And Val(left(strLayout, 1)) <= 5 And Val(Right(strLayout, 1)) <= 5 Then
                
                '更改图像分格
                Call funChangeFormat(Me.VScro.Value, Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text, _
                    mintSelectedViewer, Val(Right(strLayout, 1)), Val(left(strLayout, 1)))
                    
                '重新显示这一页图像
                Call subShowOnePage(Me.VScro.Value)
            End If
            '重新读取打印后清空参数
            mblnClearAfterPrint = IIf(GetSetting("ZLSOFT", "公共模块\zlPacsCore", "打印后清空", "1") = 0, False, True)
    Case ID_frmFilm_UnDivide            ''测试分割结果
        Dim imgTemp As DicomImage
        Dim arrImageSize() As ImageSize
        
        Set clsTruePrinter = funFillPrinterParams(False)
        
        If clsTruePrinter Is Nothing Then Exit Sub
        
        '计算每一张图的最大分辨率
        Call subCalImageMaxSize(clsTruePrinter.strFilmSize, clsTruePrinter.strFormat, clsTruePrinter.intImageResolution, arrImageSize)
        
        If UBound(arrImageSize) >= mintSelectedViewer Then
            Set imgTemp = funAssembleImage(FilmViewer(mintSelectedViewer), arrImageSize(mintSelectedViewer).intWidth, arrImageSize(mintSelectedViewer).intHeight)
        Else
            Set imgTemp = funAssembleImage(FilmViewer(mintSelectedViewer))
        End If
        If imgTemp Is Nothing Then Exit Sub
        
        '添加图像
        imgsPrint.Add imgTemp
        '图像增加了，调整页数
        Call subRecalPages
        
        '如果新添加的图像，在当前页，则重新显示这一页的图像
        If imgsPrint.Count < funGetStartImgNo(Me.VScro.Value, 1, 1) + marrPages(Me.VScro.Value).intImageCount Then
            Call subShowPrintImages(Me.VScro.Value)
        End If
        
    Case ID_frmFilm_ImgIncrease     ''图像排序，正序
        Call subImageSort(True)
    Case ID_frmFilm_ImgDecrease     ''图像排序，逆序
        Call subImageSort(False)
    End Select
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub subCheckToolBar(control As CommandBarControl)
    '切换控件选中状态
    Dim blnChecked As Boolean
    
    On Error Resume Next
    
    '如果是从裁剪状态退出，需要处理裁剪框
    If CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut, , True).Checked = True Then
        If mintCutOutViewer <> 0 And mintCutOutImage <> 0 And mintCutOutLabel <> 0 Then
            If mintCutOutViewer < FilmViewer.Count Then
                If mintCutOutImage <= FilmViewer(mintCutOutViewer).Images.Count Then
                    If mintCutOutLabel = FilmViewer(mintCutOutViewer).Images(mintCutOutImage).Labels.Count Then
                        '删除框选用的临时标注
                        FilmViewer(mintCutOutViewer).Images(mintCutOutImage).Labels.Remove mintCutOutLabel
                        Set mdcmSelectLabel = Nothing
                        FilmViewer(mintCutOutViewer).Refresh
                        '把修改过的图像上传到原始图集合中
                        Call subReloadImgsPrint
                    End If
                End If
            End If
        End If
    End If
    
    If Not control Is Nothing Then blnChecked = control.Checked
    
    CommBar_Film.Item(3).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_Pan, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_Zoom, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_RectZoom, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut, , True).Checked = False
    
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_Label_R, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_Label_L, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_Label_A, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_Label_P, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_Label_I, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_Label_S, , True).Checked = False
    
    If Not control Is Nothing Then
        CommBar_Film.Item(3).FindControl(, control.Id, , True).Checked = Not blnChecked
    End If
    
    mintCutOutViewer = 0
    mintCutOutImage = 0
    mintCutOutLabel = 0
End Sub

Private Sub CommBar_Film_GetClientBordersWidth(left As Long, top As Long, Right As Long, Bottom As Long)
    If sbStatusBar.Visible Then
        Bottom = sbStatusBar.height
    End If
End Sub

Private Sub CommBar_Film_Resize()
    On Error Resume Next
    
    Dim left As Long
    Dim top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    Me.CommBar_Film.GetClientRect left, top, Right, Bottom
    If Right >= left And Bottom >= top Then
        picBak.Move left, top, Right - left, Bottom - top
    Else
        picBak.Move 0, 0, 0, 0
    End If
End Sub

Private Sub CommBar_Film_Update(ByVal control As XtremeCommandBars.ICommandBarControl)

    '根据鼠标状态更新操作提示,并设置鼠标状态
    '鼠标的状态：0－无；1－调窗；2－漫游；3－缩放;4-无;5-框选缩放;6-裁剪:7-文字标注
    Select Case intMouseState
        Case 0
            Me.MousePointer = 0
            sbStatusBar.Panels(2).Text = "拖拽图像可以前后移动图像，把图像放到空白区域为复制图像"
        Case 1
            Me.MouseIcon = ImageListMouse.ListImages("调窗").Picture
            Me.MousePointer = 99
            sbStatusBar.Panels(2).Text = "拖拽鼠标，进入手动控制的图象窗宽窗位调节模式"
        Case 2
            Me.MouseIcon = ImageListMouse.ListImages("漫游").Picture
            Me.MousePointer = 99
            sbStatusBar.Panels(2).Text = "拖拽鼠标，在观察区内移动图象的位置，以便于更好地观察"
        Case 3
            Me.MouseIcon = ImageListMouse.ListImages("缩放").Picture
            Me.MousePointer = 99
            sbStatusBar.Panels(2).Text = "拖拽鼠标，在观察区内缩小或放大图像"
        Case 5
            Me.MouseIcon = ImageListMouse.ListImages("框选缩放").Picture
            Me.MousePointer = 99
            sbStatusBar.Panels(2).Text = "拖拽鼠标，框选出需要放大的区域，松开鼠标进行缩放"
        Case 6
            Me.MouseIcon = ImageListMouse.ListImages("裁剪").Picture
            '因为裁剪的过程中，移动裁剪框时，使用了四个方向的鼠标指针
            If Me.MousePointer = 0 Then Me.MousePointer = 99
            sbStatusBar.Panels(2).Text = "拖拽鼠标，框选出需要裁剪的区域，双击图像进行裁剪"
        Case 7
            Me.MouseIcon = ImageListMouse.ListImages(pstrSideMarker).Picture
            Me.MousePointer = 99
            sbStatusBar.Panels(2).Text = "单击鼠标，标注选中体位文字"
    End Select
    
    '更新菜单显示状态
    Select Case control.Id
        Case ID_frmFilm_TakePictures
            '没有图像的时候，照相按钮要灰掉
            control.Enabled = IIf(imgsPrint.Count >= 1, True, False)
            If control.Enabled = True Then control.Enabled = Not mblnPrinting
        Case ID_frmFilm_Label
            If pstrSideMarker = "" And control.Caption <> "标注" Then
                control.Caption = "标注"
                control.SetFocus
            ElseIf control.Caption <> pstrSideMarker And pstrSideMarker <> "" Then
                control.Caption = pstrSideMarker
                control.SetFocus
            End If
        Case ID_frmFilm_UnDivide
            '合并测试按钮，正确输入密码,进入测试状态后启动
            control.Visible = mblnInTest
    End Select
    
    '更改滚动条显示
    If VScro.Max > 1 Then
        VScro.Visible = True
    Else
        VScro.Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '输入密码（zl9PacsWork test）,则显示“合并测试”按钮
    Static strPass As String
    
    TimePass.Enabled = False
    If KeyCode = vbKeyF12 And Shift = 7 Then
        strPass = ""
        Exit Sub
    End If
    
    If KeyCode = vbKeyEscape Then
        Call subCheckToolBar(Nothing)
        intMouseState = 0
    End If
    
    If KeyCode <> vbKeyReturn Then
        '记录当前输入的字符
        If InStr(1, "1234567890 ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyCode))) <> 0 Then
            strPass = strPass & UCase(Chr(KeyCode))
        End If
        
        '输入的字符=密码，则显示合并测试按钮
        If strPass = "ZL9PACSWORK TEST" Then
            mblnInTest = True
        Else
            mblnInTest = False
        End If
    End If
    
    TimePass.Enabled = True
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strSQL As String
    
    mintFilmWidth = 11
    mintFilmHeight = 14
    mblnIsPortrait = True
    mblnIsRow = True
    mintPrintFilmCount = 0
    
    '初始化鼠标状态，根据主窗体的鼠标左键功能设置相同的鼠标初始状态
    '记录鼠标的状态：0－无；1－调窗；2－漫游；3－缩放;4-无;5-框选缩放;6-裁剪:7-文字标注，跟frmFilmView窗体有交互操作
    '调窗102，漫游103，缩放104，
    If cMouseUsage("102").lngMouseKey = 1 And Button_miWidthLevel Then  '调窗
        intMouseState = 1
    ElseIf cMouseUsage("103").lngMouseKey = 1 And Button_miCruise Then  '漫游
        intMouseState = 2
    ElseIf cMouseUsage("104").lngMouseKey = 1 And Button_miZoom Then  '缩放
        intMouseState = 3
    Else
        intMouseState = 0
    End If
    
    mintCutOutViewer = 0
    mintCutOutImage = 0
    mintCutOutLabel = 0
    mdblCutOutRatio = 0
    mblnPrinting = False
    mblnInTest = False
    mblnPrinted = False      '默认没有被打印过
    
    '读取窗体位置
    Call RestoreWinState(Me, App.ProductName)
    
    '读取工具栏位置
    mintTBMainPosition = Val(GetSetting("ZLSOFT", "公共模块\" & App.EXEName & "\frmFilm", "主工具栏位置", 0))
    If mintTBMainPosition < 0 Or mintTBMainPosition > 3 Then mintTBMainPosition = 0
    mintTBImageProcessPosition = Val(GetSetting("ZLSOFT", "公共模块\" & App.EXEName & "\frmFilm", "图像操作工具栏位置", 2))
    If mintTBImageProcessPosition < 0 Or mintTBImageProcessPosition > 3 Then mintTBImageProcessPosition = 0
    
    '创建菜单
    Call CreateBar
    '设置状态栏图标
    'Set sbStatusBar.Panels(1).Picture = f.ImgList24.ListImages("中联图标").Picture
    sbStatusBar.Panels(2).Text = "操作提示"
    sbStatusBar.Panels(3).Text = "页数："
    
    '增加1  黄捷
    '设置菜单中下拉列表项的内容
    With Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, True)
        If blLocalRun = True Then
            strSQL = "SELECT 规格标识 as 名称 FROM 影像胶片规格"
            Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
        Else
            strSQL = "SELECT 名称 FROM 影像胶片规格"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        End If
    
        While Not rsTemp.EOF
            .AddItem rsTemp!名称
            rsTemp.MoveNext
        Wend
        If .ListCount > 1 Then
            .ListIndex = 1
            CommBar_Film_Execute Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, True)
        End If
    End With
    
    With Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True)
        If blLocalRun = True Then
            strSQL = "SELECT 格式标识 as 名称 FROM 影像打印格式"
            Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
        Else
            strSQL = "SELECT 名称 FROM 影像打印格式"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        End If
        While Not rsTemp.EOF
            .AddItem rsTemp!名称
            rsTemp.MoveNext
        Wend
    End With
    
    With Me.CommBar_Film.FindControl(, ID_frmFilm_Camera, , True)
        If cDICOMPrinter.Count > 0 Then
            For i = 1 To cDICOMPrinter.Count
                .AddItem cDICOMPrinter(i).strname
            Next
            .ListIndex = 0
        End If
    End With
    
    mintCellSpacing = 35

    Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, , True).Text = GetSetting("ZLSOFT", "公共模块\" & App.EXEName & "\frmFilm", "FilmSize", "14INX17IN")
    Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text = GetSetting("ZLSOFT", "公共模块\" & App.EXEName & "\frmFilm", "FilmFormat", "STANDARD\1,1")
    Me.CommBar_Film.FindControl(, ID_frmFilm_Camera, , True).Text = GetSetting("ZLSOFT", "公共模块\" & App.EXEName & "\frmFilm", "PrinterName", "")
    Me.CommBar_Film.FindControl(, ID_frmFilm_FilmCol, , True).Checked = IIf(GetSetting("ZLSOFT", "公共模块\" & App.EXEName & "\frmFilm", "FilmPortrait", True) = "True", True, False)
    Me.CommBar_Film.FindControl(, ID_frmFilm_FilmRow, , True).Checked = Not Me.CommBar_Film.FindControl(, ID_frmFilm_FilmCol, , True).Checked
    mblnIsPortrait = Me.CommBar_Film.FindControl(, ID_frmFilm_FilmCol, , True).Checked
    
    '初始化当前的页面布局
    Call InitPageFormat(Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text)
    
    '显示第一页，空页
    Call subShowOnePage(1)
    
    mblnCheckPrinter = IIf(GetSetting("ZLSOFT", "公共模块\zlPacsCore", "检查打印机状态", "0") = 1, True, False)
    
    mblnClearAfterPrint = IIf(GetSetting("ZLSOFT", "公共模块\zlPacsCore", "打印后清空", "1") = 0, False, True)
    
    Me.VScro.Min = 1
    Me.VScro.Value = 1
     
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    mblnBegin = True
    
End Sub

Private Sub subLoadViewer(intPage As Integer)
'------------------------------------------------
'功能：加载一页Viewer，卸载多余的Viewer，并根据intViewerCount的数量重新装载Viewer。
'参数： intPage --- 需要显示的页面
'返回：无
'------------------------------------------------
    Dim i As Integer
    Dim intViewerCount As Integer
    
    If intPage > UBound(marrPages) Then Exit Sub
    
    On Error GoTo err
    
    intViewerCount = marrPages(intPage).intViewerCount
    
    '卸载多余的viewer
    For i = intViewerCount + 1 To FilmViewer.Count - 1
        Unload FilmViewer(i)
    Next
    
    '重新装载缺少的viewer
    For i = FilmViewer.Count To intViewerCount
        load FilmViewer(i)
        FilmViewer(i).Visible = True
        FilmViewer(i).CellSpacing = 2
    Next
    
    '清除原来的图像，设置每个Viewer的图像组合情况
    For i = 1 To FilmViewer.Count - 1
        FilmViewer(i).Images.Clear
        FilmViewer(i).MultiColumns = marrPages(intPage).ViewerLayout(i).intColumns
        FilmViewer(i).MultiRows = marrPages(intPage).ViewerLayout(i).intRows
    Next i
    
    '重新设置行列参数
    Call subFillPageRCCount(marrPages(intPage).strPageFormat)
    
    '重新摆放Viewer
    Call picBak_Resize
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '只有关闭主窗体的时候，才卸载胶片打印窗体，其他情况都只是隐藏胶片打印窗体
    If UnloadMode <> vbFormOwner Then
        Cancel = 1
        Me.Hide
        f.blnPrintFilm = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ''''''''''''''''''''''解除鼠标滚轮'''''''''''''''''''''''''''''''''
     '    卸载hook
    Call FilmUnhook(Me.hwnd, plngFilmPreWndProc)
    
    '保存照相参数
    SaveSetting "ZLSOFT", "公共模块\" & App.EXEName & "\frmFilm", "FilmSize", Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, , True).Text
    SaveSetting "ZLSOFT", "公共模块\" & App.EXEName & "\frmFilm", "FilmFormat", Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text
    SaveSetting "ZLSOFT", "公共模块\" & App.EXEName & "\frmFilm", "PrinterName", Me.CommBar_Film.FindControl(, ID_frmFilm_Camera, , True).Text
    SaveSetting "ZLSOFT", "公共模块\" & App.EXEName & "\frmFilm", "FilmPortrait", Me.CommBar_Film.FindControl(, ID_frmFilm_FilmCol, , True).Checked
    
    '保存窗体位置
    Call SaveWinState(Me, App.ProductName)
    
    '保存工具栏位置
    SaveSetting "ZLSOFT", "公共模块\" & App.EXEName & "\frmFilm", "主工具栏位置", Me.CommBar_Film(2).Position
    SaveSetting "ZLSOFT", "公共模块\" & App.EXEName & "\frmFilm", "图像操作工具栏位置", Me.CommBar_Film(3).Position
    
    f.blnPrintFilm = False
    mblnBegin = False
    imgsPrint.Clear
    ReDim marrPages(0)
    
    
End Sub


Private Sub mfrmFilmView_AfterClose(dcmImage As DicomObjects.DicomImage, intViewerIndex As Integer, intImageIndex As Integer)
    '关闭图像处理窗口，把处理好的图像恢复到当前胶片预览中
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    On Error GoTo err
    If intViewerIndex = 0 Or intImageIndex = 0 Then Exit Sub
    If FilmViewer.Count < intViewerIndex Then Exit Sub
    If FilmViewer(intViewerIndex).Images.Count < intImageIndex Then Exit Sub
        
    FilmViewer(intViewerIndex).Images.Remove (intImageIndex)
    FilmViewer(intViewerIndex).Images.Add dcmImage
    Call FilmViewer(intViewerIndex).Images.Move(FilmViewer(intViewerIndex).Images.Count, intImageIndex)
    
    '调整图像的位置和缩放比例
    If dcmImage.StretchToFit = False Then
        lngWidth = mfrmFilmView.dcmViewer.width / mfrmFilmView.dcmViewer.MultiColumns
        lngHeight = mfrmFilmView.dcmViewer.height / mfrmFilmView.dcmViewer.MultiRows
        Call subScaleImage(FilmViewer(intViewerIndex).Images(intImageIndex), FilmViewer(intViewerIndex), lngWidth, lngHeight)
    End If
    
    '卸载mfrmFilmView对象
    Set mfrmFilmView = Nothing
    
    '清空裁剪标记
    mintCutOutViewer = 0
    mintCutOutImage = 0
    mintCutOutLabel = 0
    
    '把更新提交会原始图像集
    '把标注过的图像上传到原始图集合中
    Call subReloadImgsPrint
    
    '同步图像
    mintSelectedViewer = intViewerIndex
    mintSelectedImage = intImageIndex
    Set SelectedImage = FilmViewer(intViewerIndex).Images(intImageIndex)
    Call subSynchronalImg(False, IMG_SYN_All)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picBak_Resize()

'这是一个单纯调整Viewer位置的过程，不涉及对Viewer的加载，图像的加载等

    Dim i As Integer, j As Integer, k As Integer
    Dim l As Integer, t As Integer, w As Integer, h As Integer
    
    If Not mblnBegin Then Exit Sub
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '先调整滚动条的位置
    VScro.Move Me.picBak.ScaleWidth - VScro.width, 0, VScro.width, Me.picBak.ScaleHeight
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '计算胶片和屏幕的长宽，或者宽长比例
    mdubFilmRate = IIf(mblnIsPortrait, mintFilmWidth / mintFilmHeight, mintFilmHeight / mintFilmWidth)
    mdubScreenRate = (picBak.ScaleWidth - Me.VScro.width) / picBak.ScaleHeight
  
    '隐藏背景的picFilm
    picFilm.Visible = False
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '设置picFilm的位置
    If mdubFilmRate < mdubScreenRate Then
        picFilm.top = 0
        picFilm.height = picBak.ScaleHeight
        picFilm.width = picFilm.height * mdubFilmRate '- Me.VScro.width
        picFilm.left = Abs(picBak.ScaleWidth - picFilm.width - 250) / 2
    Else
        picFilm.left = 0
        picFilm.width = Abs(picBak.ScaleWidth - Me.VScro.width)
        picFilm.height = picFilm.width / mdubFilmRate
        picFilm.top = Abs(picBak.ScaleHeight - picFilm.height) / 2
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '摆放Viewer的位置
    k = 1
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To UBound(marrRCCount)
        If mblnIsRow Then
            For j = 1 To marrRCCount(i)
                h = Abs(picFilm.ScaleHeight / UBound(marrRCCount) - mintCellSpacing * 2)
                w = Abs(picFilm.ScaleWidth / marrRCCount(i) - mintCellSpacing * 2)
                l = Abs(picFilm.ScaleWidth / marrRCCount(i) * (j - 1) + mintCellSpacing)
                t = Abs(picFilm.ScaleHeight / UBound(marrRCCount) * (i - 1) + mintCellSpacing)
                If Me.CommBar_Film.FindControl(, ID_frmFilm_RectPhotCase, , True).Checked Then
                    If h > w Then
                         t = t + (h - w) / 2
                         h = w
                    Else
                         l = l + (w - h) / 2
                         w = h
                    End If
                End If
                FilmViewer(k).Move l, t, w, h
                k = k + 1
            Next
        Else
            For j = 1 To marrRCCount(i)
                h = Abs(picFilm.ScaleHeight / marrRCCount(i) - mintCellSpacing * 2)
                w = Abs(picFilm.ScaleWidth / UBound(marrRCCount) - mintCellSpacing * 2)
                l = Abs(picFilm.ScaleWidth / UBound(marrRCCount) * (i - 1) + mintCellSpacing)
                t = Abs(picFilm.ScaleHeight / marrRCCount(i) * (j - 1) + mintCellSpacing)
                If Me.CommBar_Film.FindControl(, ID_frmFilm_RectPhotCase, , True).Checked Then
                    If h > w Then
                         t = t + (h - w) / 2
                         h = w
                    Else
                         l = l + (w - h) / 2
                         w = h
                    End If
                End If
                FilmViewer(k).Move l, t, w, h
                k = k + 1
            Next
        End If
    Next
    
    For i = 1 To FilmViewer.Count - 1
        '显示图像选择框
        Call subReScaleViewerFrame(FilmViewer(i))
        FilmViewer(i).Visible = True
    Next i
    
    picFilm.Visible = True
End Sub

Private Sub subLoadPrintImage(intPage As Integer)
'------------------------------------------------
'功能：显示一页图像，向Viewer中顺序装入imgsPrint中存储的图像。
'参数： intPage -- 显示的页数
'返回：无
'------------------------------------------------
    Dim i As Integer, j As Integer
    Dim intStart As Integer
    Dim blnUpLoad As Boolean
    
    On Error GoTo err
    
    For i = 1 To marrPages(intPage).intViewerCount
        FilmViewer(i).Images.Clear
        
        '重新显示图像选择框
        Call subFilmDispframe(FilmViewer(i))
    Next i
    
    '循环每一个Viewer，添加图像
    intStart = funGetStartImgNo(intPage, 1, 1)
    For i = 1 To marrPages(intPage).intViewerCount
        
        If intStart > imgsPrint.Count Then Exit For
        
        '从正本中添加图像到Viewer
        For j = 1 To FilmViewer(i).MultiColumns * FilmViewer(i).MultiRows
            If intStart > imgsPrint.Count Then Exit For
            '添加图像
            FilmViewer(i).Images.Add imgsPrint(intStart)
            
            '修正图像的位置和缩放比例
            If FilmViewer(i).Images(j).StretchToFit = False Then
                If Val(funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_V宽度)) <> 0 And Val(funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_V高度)) <> 0 Then
                    Call subScaleImage(FilmViewer(i).Images(j), FilmViewer(i), _
                        Val(funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_V宽度)), _
                        Val(funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_V高度)))
                Else
                    '说明有图像是最近添加的，TAG中没有记录Viewer的宽度和高度，需要添加，最后还要reload到后台图像集中
                    Call funSetTagVal(FilmViewer(i).Images(j), TAG_V宽度, CStr(FilmViewer(i).width / FilmViewer(i).MultiColumns))
                    Call funSetTagVal(FilmViewer(i).Images(j), TAG_V高度, CStr(FilmViewer(i).height / FilmViewer(i).MultiRows))
                    blnUpLoad = True
                End If
            End If
            
            '设置图像选择标记
            If funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_选择) = "Select" Then
                Call subImageSelect(i, j, True)
            End If
                
            intStart = intStart + 1
        Next j
        
    Next i
    
    '设置被当前和被选中的图像
    If FilmViewer.Count > 1 And FilmViewer(1).Images.Count > 0 Then
        Call subImageCurrent(1, 1, True)
    Else
        mintSelectedViewer = 0
        mintSelectedImage = 0
        Set SelectedImage = Nothing
        mblnPrinted = False      '没有图像，打印标记设置成默认值False
    End If
    
    '将图像信息回传到正本中
    If blnUpLoad Then
        Call subReloadImgsPrint
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub TimePass_Timer()
    Call Form_KeyDown(vbKeyF12, 7)   '清除静态变量
End Sub

Private Sub FilmViewer_DblClick(Index As Integer)
    
    '在图像上双击时，有以下用法：
    '1、打开图像处理窗口
    '2、进行图像裁剪
    '使用裁剪标记区别两种用法
    If mintCutOutViewer = 0 Or mintCutOutImage = 0 Or mintCutOutLabel = 0 Then
        '打开图像处理窗口
        Call subOpenFilmView
        
    Else    '图像裁剪
        If mintCutOutViewer >= FilmViewer.Count Then Exit Sub
        If mintCutOutImage > FilmViewer(mintCutOutViewer).Images.Count Then Exit Sub
        If mintCutOutLabel <> FilmViewer(mintCutOutViewer).Images(mintCutOutImage).Labels.Count Then Exit Sub
        
        Dim Image As DicomImage
        Dim i As Integer
        Dim lblTemp As DicomLabel
        Dim sourceImage As DicomImage
        
        Set sourceImage = FilmViewer(mintCutOutViewer).Images(mintCutOutImage)
        Set Image = CutOutAImage(sourceImage)
        
        Image.Name = "ZLPIC"
        '删除框选用的临时标注
        sourceImage.Labels.Remove mintCutOutLabel
        Set mdcmSelectLabel = Nothing
        
        Call subWriteDicomPara(sourceImage, Image)
        
        '把原来图像的标注，添加到现在的图像中
        Image.Labels.Clear
        For i = 1 To sourceImage.Labels.Count
            Image.Labels.Add sourceImage.Labels(i)
        Next i
        
        '把新生成的图像，添加到Viewer中
        If mintCutOutImage = 1 And FilmViewer(mintCutOutViewer).Images.Count = 1 Then
            FilmViewer(mintCutOutViewer).Images.Clear
            FilmViewer(mintCutOutViewer).Images.Add Image
        Else
            FilmViewer(mintCutOutViewer).Images.Remove mintCutOutImage
            FilmViewer(mintCutOutViewer).Images.Add Image
            FilmViewer(mintCutOutViewer).Images.Move FilmViewer(mintCutOutViewer).Images.Count, mintCutOutImage
        End If
        
        '图像放入Viewer中后，重新显示标尺，这个时候标尺和单位才是准确的
        Call UpdateRuler(Image, True)
        
        mintCutOutViewer = 0
        mintCutOutImage = 0
        mintCutOutLabel = 0
        Me.MousePointer = vbArrow
        
        '提交正本图像
        Call subReloadImgsPrint
    End If
End Sub

Private Sub FilmViewer_DragDrop(Index As Integer, Source As control, x As Single, y As Single)
    '放下拖拽的图像
    Dim intOldImgIndex As Integer
    Dim intOldViewerIndex As Integer
    Dim intOldImgsPrintIndex As Integer
    Dim intNewImgIndex As Integer
    Dim intImgIndex As Integer
    Dim img As DicomImage
    Dim i As Integer
    Dim j As Integer
    Dim intOldDelta As Integer
    Dim blnNewMoveUp As Boolean
    
    On Error GoTo err
    
    If Source.Name = "FilmViewer" And Source.Images.Count > 0 Then
        '提取图像的旧位置
        intOldImgIndex = Val(Source.Tag)
        If intOldImgIndex <= 0 Then Exit Sub
        
        intOldViewerIndex = Val(Source.Index)
        If intOldViewerIndex <= 0 Or intOldViewerIndex > FilmViewer.Count Then Exit Sub
        
        '提取旧图像在正本中的位置
        intOldImgsPrintIndex = funGetStartImgNo(VScro.Value, intOldViewerIndex, intOldImgIndex)
        If intOldImgsPrintIndex <= 0 Or intOldImgsPrintIndex > imgsPrint.Count Then
            Exit Sub
        End If
                            
        '提取图像的新位置
        intImgIndex = FilmViewer(Index).ImageIndex(x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
        intNewImgIndex = funGetStartImgNo(VScro.Value, Index, intImgIndex)
        
        If intOldImgsPrintIndex = intNewImgIndex And intImgIndex <> 0 Then Exit Sub
        
        '开始摆放图像之前，先把原来图像的TAG提交到图像正本中
        Call subReloadImgsPrint
        
        '重新摆放图像
        If intImgIndex = 0 Then     '说明是拖到了空的地方，需要往正本增加一个图
            '检查原图是否被选中，如果是被选中的，处理多选移动
            If funGetTagVal(FilmViewer(intOldViewerIndex).Images(intOldImgIndex).Tag, TAG_选择) = "Select" Then
                '当前图像被选中，处理多选移动
                For i = 1 To FilmViewer.Count - 1
                    For j = 1 To FilmViewer(i).Images.Count
                        If funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_选择) = "Select" Then
                            Set img = New DicomImage
                            '提取旧图像在正本中的位置
                            intOldImgsPrintIndex = funGetStartImgNo(VScro.Value, i, j)
                            If intOldImgsPrintIndex <= 0 Or intOldImgsPrintIndex > imgsPrint.Count Then
                                Exit For
                            End If
                            
                            Set img = imgsPrint(intOldImgsPrintIndex)
                            imgsPrint.Add img
                        End If
                    Next j
                Next i
            Else
                '只处理一个图
                Set img = New DicomImage
                Set img = imgsPrint(intOldImgsPrintIndex)
                imgsPrint.Add img
            End If
        Else    '交换图像位置
            '如果往后移，新位置要-1
            If intNewImgIndex > intOldImgsPrintIndex Then
'                intNewImgIndex = intNewImgIndex - 1
                blnNewMoveUp = True
            Else
                blnNewMoveUp = False
            End If
            If intNewImgIndex <= 0 Or intNewImgIndex > imgsPrint.Count Then Exit Sub
            
            
            '检查原图是否被选中，如果是被选中的，处理多选移动
            If funGetTagVal(FilmViewer(intOldViewerIndex).Images(intOldImgIndex).Tag, TAG_选择) = "Select" Then
                '当前图像被选中，处理多选移动
                intOldDelta = 0
                For i = 1 To FilmViewer.Count - 1
                    For j = 1 To FilmViewer(i).Images.Count
                        If funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_选择) = "Select" Then
                            '按照图像原来的顺序移动被选中的图像
                            
                            '提取旧图像在正本中的位置
                            intOldImgsPrintIndex = funGetStartImgNo(VScro.Value, i, j)
                            
                            If intOldImgsPrintIndex < intNewImgIndex Then
                                intOldImgsPrintIndex = intOldImgsPrintIndex + intOldDelta
                            End If
                            
                            If intOldImgsPrintIndex <= 0 Or intOldImgsPrintIndex > imgsPrint.Count Then
                                Exit For
                            End If
                            
                            '如果往前移，但是之前已经有了往后移的效果，则需要+1
                            If (intNewImgIndex < intOldImgsPrintIndex) And blnNewMoveUp Then
                                intNewImgIndex = intNewImgIndex + 1
                                blnNewMoveUp = False
                            ElseIf (intNewImgIndex > intOldImgsPrintIndex) And blnNewMoveUp = False Then
                                intNewImgIndex = intNewImgIndex - 1
                                blnNewMoveUp = True
                            End If
                            
                            If intNewImgIndex <> intOldImgsPrintIndex Then
                                '移动图像
                                Call imgsPrint.Move(intOldImgsPrintIndex, intNewImgIndex)
                                If intOldImgsPrintIndex > intNewImgIndex Then
                                    intNewImgIndex = intNewImgIndex + 1
                                Else
                                    intOldDelta = intOldDelta - 1
                                End If
                            End If
                        End If
                    Next j
                Next i
            Else
                '移动图像
                Call imgsPrint.Move(intOldImgsPrintIndex, intNewImgIndex)
            End If
        End If
        '重新显示图像
        Call subShowPrintImages(Me.VScro.Value)
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FilmViewer_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '处理Del
    If KeyCode = 46 Then        'Delete
        Call subDelImage
    End If
End Sub

Private Sub FilmViewer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim intImgIndex As Integer
    Dim ls As DicomLabels
    
    On Error GoTo err
    
    mintBaseX = x
    mintBaseY = y
    
    '切换了图像
    If Index >= FilmViewer.Count Then Exit Sub
    intImgIndex = FilmViewer(Index).ImageIndex(x, y)
    If FilmViewer(Index).Images.Count > 0 And intImgIndex <> 0 Then
        '切换图像,先恢复原来图像的选择框
        If mintSelectedViewer > 0 And mintSelectedViewer < FilmViewer.Count Then
            If mintSelectedImage > 0 And mintSelectedImage <= FilmViewer(mintSelectedViewer).Images.Count Then
                If funGetTagVal(FilmViewer(mintSelectedViewer).Images(mintSelectedImage).Tag, TAG_选择) = "Select" Then
                    Call subImageSelect(mintSelectedViewer, mintSelectedImage, True)
                Else
                    Call subImageSelect(mintSelectedViewer, mintSelectedImage, False)
                End If
                FilmViewer(mintSelectedViewer).Refresh
            End If
        End If
        
        '再设置当前图像的选择框
        Call subImageCurrent(Index, intImgIndex, True)
        
        '设置窗宽窗位弹出菜单
        Call subSetWidthLevelF(SelectedImage, Me)
        
        If Button = 1 Then
            'intMouseState 鼠标的状态：0－无；1－调窗；2－漫游；3－缩放;4-无;5-框选缩放;6-裁剪:7-文字标注
            If intMouseState = 6 Then '裁剪
                '裁剪状态下的鼠标down，有三种操作：1、画裁剪框（记录标记）；2、移动裁剪框(有焦点) ；3、双击进行裁剪
                If mintCutOutViewer = 0 Or mintCutOutImage = 0 Or mintCutOutLabel = 0 Then  '画裁剪框
                    '判断是固定裁剪，还是自由裁剪
                    If (CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut, , True).IconId <> ID_frmFilm_CutOut_Custom) _
                        And (CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut, , True).IconId <> ID_frmFilm_CutOut) Then
                        '固定裁剪
                        Call subCutOutRatio(CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut, , True).IconId)
                    Else
                        '增加框选标注
                        FilmViewer(Index).Images(intImgIndex).Labels.Add GetNewLabel(doLabelRectangle, FilmViewer(Index).ImageXPosition(x, y), FilmViewer(Index).ImageYPosition(x, y), 0, 0)
                        Set mdcmSelectLabel = FilmViewer(Index).Images(intImgIndex).Labels(FilmViewer(Index).Images(intImgIndex).Labels.Count)
                        mdcmSelectLabel.Tag = CUT_LABEL
                        mblnDcmViewDown = True
                        mintCutOutViewer = Index
                        mintCutOutImage = intImgIndex
                        mintCutOutLabel = FilmViewer(Index).Images(intImgIndex).Labels.Count
                    End If
                Else            '开始移动裁剪框
                    Set ls = FilmViewer(Index).LabelHits(x, y, False, False, True)
                    If ls.Count <> 0 And Me.MousePointer <> vbArrow Then
                        '开始移动裁剪框
                        If ls(1).Tag = CUT_LABEL And SelectedImage.Labels(SelectedImage.Labels.Count).Tag = CUT_LABEL Then
                            mblnLabelMoving = True
                        End If
                    End If
                End If
            End If
            If intMouseState = 5 Then      '框选缩放
                '增加框选标注
                FilmViewer(Index).Images(intImgIndex).Labels.Add GetNewLabel(doLabelRectangle, FilmViewer(Index).ImageXPosition(x, y), FilmViewer(Index).ImageYPosition(x, y), 0, 0)
                Set mdcmSelectLabel = FilmViewer(Index).Images(intImgIndex).Labels(FilmViewer(Index).Images(intImgIndex).Labels.Count)
                mblnDcmViewDown = True
            End If
            If intMouseState = 7 Then       '文字标注
                Dim dcmLabel As DicomLabel
                Set dcmLabel = GetNewLabel(doLabelText, FilmViewer(Index).ImageXPosition(x, y), FilmViewer(Index).ImageYPosition(x, y), 0, 0)
                FilmViewer(Index).Images(intImgIndex).Labels.Add dcmLabel
                dcmLabel.AutoSize = True
                dcmLabel.Margin = 0
                dcmLabel.Text = pstrSideMarker
                dcmLabel.Shadow = doShadowAll
                dcmLabel.ShowTextBox = True
                dcmLabel.Font.Bold = True
                dcmLabel.Tag = POSTURE_LABEL
                intMouseState = 0
                pstrSideMarker = ""
                '把标注过的图像上传到原始图集合中
                Call subReloadImgsPrint
            End If
            'Ctrl单个选择
            If Shift = 2 Then
                If funGetTagVal(FilmViewer(Index).Images(intImgIndex).Tag, TAG_选择) = "Select" Then
                    Call subImageSelect(mintSelectedViewer, mintSelectedImage, False)
                Else
                    Call subImageSelect(mintSelectedViewer, mintSelectedImage, True)
                End If
                
                '把选择的图像上传到原始图集合中
                Call subReloadImgsPrint
            End If
            
            'Shift连续选择
            If Shift = 1 Then
                Call subShiftSelect(Index, intImgIndex)
            End If
            
            If intMouseState = 0 Then   '鼠标无任何状态，开始拖拽
                If FilmViewer(Index).Images.Count > 0 Then
                    'tag 包含一个字段，Viewer其中图像所在的索引
                    FilmViewer(Index).Tag = FilmViewer(Index).ImageIndex(x, y)
                    FilmViewer(Index).Drag
                End If
            End If
        End If
        FilmViewer(Index).Refresh
    End If
    Exit Sub
err:
End Sub

Private Sub SelAllImage(blnSelect As Boolean)
'------------------------------------------------
'功能：全选或者全不选所有图像，设置所有图像的选择框的颜色
'       被选中的图像边框为红色，没有被选中的图像边框为白色
'参数： blnSelect -- True 选择图像；False 全清选择
'返回：无
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo err
    
    For i = 1 To FilmViewer.Count - 1
        If FilmViewer(i).Visible = True Then
            For j = 1 To FilmViewer(i).Images.Count
                Call subImageSelect(i, j, blnSelect)
            Next j
            FilmViewer(i).Refresh
        End If
    Next i
    
    '把选择的图像上传到原始图集合中
    Call subReloadImgsPrint
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub FilmViewer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim dblZoom As Double
    Dim ls As DicomLabels
    Dim lngXOffset As Long
    Dim lngYOffset As Long
    Dim lblCUT As DicomLabel
    
    On Error GoTo err
    
    If SelectedImage Is Nothing Then Exit Sub
    
    If (Button = 1 And intMouseState = 1) Or (Button = 4 And intMouseWheelDrag = 2) _
        Or (Button = 2 And cMouseUsage("102").lngMouseKey = 2 And Button_miWidthLevel) Then  '调窗
        If SelectedImage.VOILUT = 1 Then SelectedImage.VOILUT = 0
        SelectedImage.width = SelectedImage.width + (x - mintBaseX) * lngWidthLevelStep / 5
        SelectedImage.Level = SelectedImage.Level + (y - mintBaseY) * lngWidthLevelStep / 5
        SelectedImage.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & SelectedImage.width & "-L:" & SelectedImage.Level
        mintBaseX = x
        mintBaseY = y
        FilmViewer(Index).Refresh
    ElseIf (Button = 1 And intMouseState = 2) Or (Button = 4 And intMouseWheelDrag = 0) _
        Or (Button = 2 And cMouseUsage("103").lngMouseKey = 2 And Button_miCruise) Then '漫游
        subCenterZoom SelectedImage, FilmViewer(Index), SelectedImage.ActualZoom
        SelectedImage.ScrollX = SelectedImage.ScrollX - (x - mintBaseX) * lngCruiseStep / 5
        SelectedImage.ScrollY = SelectedImage.ScrollY - (y - mintBaseY) * lngCruiseStep / 5
        mintBaseX = x
        mintBaseY = y
    ElseIf (Button = 1 And intMouseState = 3) Or (Button = 4 And intMouseWheelDrag = 1) _
        Or (Button = 2 And cMouseUsage("104").lngMouseKey = 2 And Button_miZoom) Then '缩放
        '缩放单位是0.01倍
        dblZoom = SelectedImage.ActualZoom * (1 + (mintBaseY - y) * lngZoomStep / 5 * 0.001)
        If dblZoom < 0.01 Then dblZoom = 0.01
        If dblZoom > 64 Then dblZoom = 64
        subCenterZoom SelectedImage, FilmViewer(Index), dblZoom
        
        If SelectedImage.Labels.Count >= G_INT_SYS_LABEL_RULLER Then
            If SelectedImage.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then  '更新标尺单位
                UpdateRuler SelectedImage, True
            End If
        End If
        
        mintBaseX = x
        mintBaseY = y
    ElseIf Button = 1 And (intMouseState = 5 Or intMouseState = 6) Then  '框选缩放
        If mblnDcmViewDown = True Then
            mdcmSelectLabel.width = FilmViewer(Index).ImageXPosition(x, y) - mdcmSelectLabel.left
            mdcmSelectLabel.height = FilmViewer(Index).ImageYPosition(x, y) - mdcmSelectLabel.top
            FilmViewer(Index).Refresh
        End If
    End If
    
    If intMouseState = 6 And mintCutOutViewer <> 0 And mintCutOutImage <> 0 And mintCutOutLabel <> 0 Then
        Set ls = FilmViewer(Index).LabelHits(x, y, False, False, True)
        If Button = 1 Then          '鼠标被按下
            If mblnLabelMoving = True Then
                Call subaCorrectCursor(FilmViewer(Index), SelectedImage, x, y)
                Set lblCUT = SelectedImage.Labels(SelectedImage.Labels.Count)
                
                If (Me.MousePointer = vbSizeWE And (SelectedImage.RotateState = doRotateNormal Or SelectedImage.RotateState = doRotate180)) _
                    Or (Me.MousePointer = vbSizeNS And (SelectedImage.RotateState = doRotateLeft Or SelectedImage.RotateState = doRotateRight)) Then       '左右移动
                    
                    lngXOffset = (FilmViewer(Index).ImageXPosition(x, y) - FilmViewer(Index).ImageXPosition(mintBaseX, mintBaseY))
                    If Abs(lblCUT.left - FilmViewer(Index).ImageXPosition(x, y)) > Abs(lblCUT.left + lblCUT.width - FilmViewer(Index).ImageXPosition(x, y)) Then '右边的移动
                            lblCUT.width = lblCUT.width + lngXOffset
                    Else    '左边线移动
                            lblCUT.left = lblCUT.left + lngXOffset
                            lblCUT.width = lblCUT.width - lngXOffset
                    End If
                    If mdblCutOutRatio <> 0 Then    '保持固定比例
                        lblCUT.height = lblCUT.width / mdblCutOutRatio
                    End If
                ElseIf (Me.MousePointer = vbSizeNS And (SelectedImage.RotateState = doRotateNormal Or SelectedImage.RotateState = doRotate180)) _
                    Or (Me.MousePointer = vbSizeWE And (SelectedImage.RotateState = doRotateLeft Or SelectedImage.RotateState = doRotateRight)) Then   '上下移动
                    
                    lngYOffset = (FilmViewer(Index).ImageYPosition(x, y) - FilmViewer(Index).ImageYPosition(mintBaseX, mintBaseY))
                    If Abs(lblCUT.top - FilmViewer(Index).ImageYPosition(x, y)) > Abs(lblCUT.top + lblCUT.height - FilmViewer(Index).ImageYPosition(x, y)) Then    '下面线的移动
                        lblCUT.height = lblCUT.height + lngYOffset
                        
                    Else    '上面线移动
                        lblCUT.top = lblCUT.top + lngYOffset
                        lblCUT.height = lblCUT.height - lngYOffset
                    End If
                    If mdblCutOutRatio <> 0 Then    '保持固定比例
                        lblCUT.width = lblCUT.height * mdblCutOutRatio
                    End If
                ElseIf Me.MousePointer = vbSizePointer Then     '整体移动
                
                    lngXOffset = (FilmViewer(Index).ImageXPosition(x, y) - FilmViewer(Index).ImageXPosition(mintBaseX, mintBaseY))
                    lngYOffset = (FilmViewer(Index).ImageYPosition(x, y) - FilmViewer(Index).ImageYPosition(mintBaseX, mintBaseY))
                    lblCUT.top = lblCUT.top + lngYOffset
                    lblCUT.left = lblCUT.left + lngXOffset
                End If
                mintBaseX = x
                mintBaseY = y
                FilmViewer(Index).Refresh
            End If
        ElseIf Button = 0 Then
            If ls.Count <> 0 Then
                If Abs(ls(1).left - FilmViewer(Index).ImageXPosition(x, y)) < 4 Or Abs(ls(1).left + ls(1).width - FilmViewer(Index).ImageXPosition(x, y)) < 4 Then
                    If SelectedImage.RotateState = doRotateLeft Or SelectedImage.RotateState = doRotateRight Then
                        Me.MousePointer = vbSizeNS
                    Else
                        Me.MousePointer = vbSizeWE
                    End If
                ElseIf Abs(ls(1).top - FilmViewer(Index).ImageYPosition(x, y)) < 4 Or Abs(ls(1).top + ls(1).height - FilmViewer(Index).ImageYPosition(x, y)) < 4 Then
                    If SelectedImage.RotateState = doRotateLeft Or SelectedImage.RotateState = doRotateRight Then
                        Me.MousePointer = vbSizeWE
                    Else
                        Me.MousePointer = vbSizeNS
                    End If
                Else
                    Me.MousePointer = vbSizePointer
                End If
            Else
                Me.MousePointer = vbArrow
            End If
        End If
    End If
    Exit Sub
err:
End Sub

Private Sub FilmViewer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
    
    On Error GoTo err
    
    If Button = 1 Then
        If intMouseState <> 0 Then
            If intMouseState = 5 And mblnDcmViewDown Then    '框选缩放
                lngLeft = SelectedImage.Labels(SelectedImage.Labels.Count).left * SelectedImage.ActualZoom
                lngTop = SelectedImage.Labels(SelectedImage.Labels.Count).top * SelectedImage.ActualZoom
                lngWidth = SelectedImage.Labels(SelectedImage.Labels.Count).width * SelectedImage.ActualZoom
                lngHeight = SelectedImage.Labels(SelectedImage.Labels.Count).height * SelectedImage.ActualZoom
                
                '调整宽高
                If lngWidth < 0 Then
                    lngLeft = lngLeft + lngWidth
                    lngWidth = -lngWidth
                End If
                
                If lngHeight < 0 Then
                    lngTop = lngTop + lngHeight
                    lngHeight = -lngHeight
                End If
                
                RectangleZoom FilmViewer(Index), SelectedImage, lngLeft, lngTop, lngWidth, lngHeight
                
                '删除框选用的临时标注
                SelectedImage.Labels.Remove SelectedImage.Labels.Count
                Set mdcmSelectLabel = Nothing
                FilmViewer(Index).Refresh
            ElseIf intMouseState = 6 Then
                If mblnDcmViewDown Then       '裁剪
                    '不做任何操作
                    '如果裁剪框为0 ，则取删除裁剪框，清除裁剪的标记
                    If mdcmSelectLabel.width = 0 Or mdcmSelectLabel.height = 0 Then
                        '删除框选用的临时标注
                        SelectedImage.Labels.Remove SelectedImage.Labels.Count
                        Set mdcmSelectLabel = Nothing
                        FilmViewer(Index).Refresh
                        
                        mintCutOutViewer = 0
                        mintCutOutImage = 0
                        mintCutOutLabel = 0
                    End If
                End If
            End If
            '图像不为空
        End If
    End If
    
    '同步,''intMouseState：0－无；1－调窗；2－漫游；3－缩放;4-无;5-框选缩放;6-裁剪:7-文字标注
    If FilmViewer(Index).Images.Count > 0 Then
        If (Button = 1 And intMouseState = 1) Or (Button = 4 And intMouseWheelDrag = 2) _
            Or (Button = 2 And cMouseUsage("102").lngMouseKey = 2 And Button_miWidthLevel) Then   '调窗
            Call subSynchronalImg(False, IMG_SYN_WINDOW)
        ElseIf (Button = 1 And (intMouseState = 2 Or intMouseState = 3 Or intMouseState = 5)) _
            Or (Button = 4 And (intMouseWheelDrag = 0 Or intMouseWheelDrag = 1)) _
            Or (Button = 2 And cMouseUsage("103").lngMouseKey = 2 And Button_miCruise) _
            Or (Button = 2 And cMouseUsage("104").lngMouseKey = 2 And Button_miZoom) Then
            Call subSynchronalImg(False, IMG_SYN_ZOOMPAN)
        End If
    End If
            
    mblnDcmViewDown = False
    mblnLabelMoving = False
    Exit Sub
err:
End Sub

Private Sub VScro_Change()
    
    If Me.VScro.Value = 0 Then
        sbStatusBar.Panels(3).Text = "页数："
        Exit Sub
    End If
    
    If UBound(marrPages) = 0 Then Exit Sub
    
    '将当前页的格式设置成工具栏菜单
    If marrPages(Me.VScro.Value).strPageFormat <> "" Then
        Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text = marrPages(Me.VScro.Value).strPageFormat
        Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Caption = marrPages(Me.VScro.Value).strPageFormat
    End If
    
    '重新显示当前页的图像
    Call subShowOnePage(Me.VScro.Value)
    
    sbStatusBar.Panels(3).Text = "页数：" & VScro.Value & "/" & VScro.Max
End Sub

Private Sub subFilmDispframe(v As DicomViewer)
'------------------------------------------------
'功能：在Viewer上显示矩形框
'参数：v－－需要显示矩形框的Viewer
'返回：无
'------------------------------------------------
    Dim w As Integer, h As Integer
    Dim l As DicomLabel
    Dim i As Integer
    
    v.Labels.Clear
    For i = 1 To v.MultiColumns * v.MultiRows
        w = v.width / Screen.TwipsPerPixelX / v.MultiColumns - 2
        h = v.height / Screen.TwipsPerPixelY / v.MultiRows - 2
        Set l = New DicomLabel
        l.LabelType = 2     '矩形标注
        l.width = w
        l.height = h
        l.left = ((i - 1) Mod v.MultiColumns) * (w + 2) + 1
        l.top = ((i - 1) \ v.MultiColumns) * (h + 2) + 1
        v.Labels.Add l
    Next i
    v.Refresh
End Sub

Private Sub subReScaleViewerFrame(v As DicomViewer)
'------------------------------------------------
'功能：调整Viewer上面矩形框的位置
'参数：v－－需要显示矩形框的Viewer
'返回：无
'------------------------------------------------
    Dim w As Integer, h As Integer
    Dim l As DicomLabel
    Dim i As Integer
    
    If v.Labels.Count <> v.MultiColumns * v.MultiRows Then
        '如果矩形框的数量不对，则重新创建
        Call subFilmDispframe(v)
    Else
        For i = 1 To v.MultiColumns * v.MultiRows
            w = v.width / Screen.TwipsPerPixelX / v.MultiColumns - 2
            h = v.height / Screen.TwipsPerPixelY / v.MultiRows - 2
            Set l = v.Labels(i)
            l.LabelType = 2     '矩形标注
            l.width = w
            l.height = h
            l.left = ((i - 1) Mod v.MultiColumns) * (w + 2) + 1
            l.top = ((i - 1) \ v.MultiColumns) * (h + 2) + 1
        Next i
        v.Refresh
    End If
    
End Sub


Private Function subPrintFilm(clsOnePrinter As clsDicomPrint) As Boolean
'------------------------------------------------
'功能：往打印机发送打印信号，将图像发送给打印机。
'参数：clsOnePrinter－－记录打印机参数的类
'返回：无
'上级函数或过程：CommBar_Film_Execute
'下级函数或过程：无
'引用的外部参数：intViewerCount
'编制人：黄捷
'------------------------------------------------
    
    '图像所在地从 k = 1 To intViewerCount的 viewer(k)中
    
    '判断图像数量是否大于等于1，没有图像则直接退出
    If FilmViewer.Count <= 1 Then
        Exit Function
    End If
    
    If clsOnePrinter Is Nothing Then
        Exit Function
    End If
    
    Dim printer As New DicomPrint, Thisim As DicomImage
    Dim k As Integer, i As Integer, j As Integer
    Dim strSQL As String
    Dim StrPrintLog As String, StrPrintPage As String
    Dim strImageUIDS As String  '记录图像的实例UID，用,分隔
    Dim intCurPage As Integer   '记录当前显示的页面
    Dim arrImageSize() As ImageSize
    Dim arrTempUIDs() As String
    Dim strTempUIDs As String
    
    printer.Node = clsOnePrinter.strIPAddress
    printer.Port = clsOnePrinter.lngPort
    printer.CallingAE = clsOnePrinter.strSCUAETitle
    printer.CalledAE = clsOnePrinter.strAETitle
    intCurPage = Me.VScro.Value
    
    On Error GoTo err1
    
    
    '循环打印每一页
    strImageUIDS = ","
    For j = 1 To Me.VScro.Max
        If mintPageRange = 0 Or j = intCurPage Then
            
            StrPrintPage = "," '记录每张胶片上打印的序列UID,不记录重复的
            '首先设置FilmSession中的参数，然后再Open打印机
            
            ''''''''''FilmSession 的参数''''''''''''''''''''
            '打印份数，必须
            If clsOnePrinter.lngCopies <> 0 Then
                printer.Copies = clsOnePrinter.lngCopies
            Else
                printer.Copies = 1
            End If
            
            'Print Priority 优先级，可选
            If clsOnePrinter.strPriority <> "" Then
                printer.Session.Attributes.Add &H2000, &H20, clsOnePrinter.strPriority
            End If
            'Medium Type 介质类型，可选
            If clsOnePrinter.strMedium <> "" Then
                printer.Session.Attributes.Add &H2000, &H30, clsOnePrinter.strMedium
            End If
            'Film Destination 介质目标，可选
            If clsOnePrinter.strFilmBox <> "" Then
                printer.Session.Attributes.Add &H2000, &H40, clsOnePrinter.strFilmBox
            End If
            
            '''''''''''''''''''''''''''''''''打开打印机''''''''''''''''''''''
            'Open的时候，会把FilmSession的参数用N-CREATE的方式传给打印机
            printer.Open
            
            '设置错误处理
            On Error GoTo err2
            
            '检测打印机返回的状态
            If mblnCheckPrinter = True Then
                If Not printer.printer Is Nothing Then
                    '检查Printer的（2110，0010）Printer Status和（2110，0020）Printer Status Info
                    If printer.printer.Attributes(&H2110, &H10).Exists And Not IsNull(printer.printer.Attributes(&H2110, &H10).Value) Then
                        If printer.printer.Attributes(&H2110, &H10).Value = "WARNING" Or _
                            printer.printer.Attributes(&H2110, &H10).Value = "FAILURE" Then
                            
                            '出现警告或者错误
                            If printer.printer.Attributes(&H2110, &H20).Exists And Not IsNull(printer.printer.Attributes(&H2110, &H20).Value) Then
                                '同时返回Printer Status和 Printer Status Info的警告或者错误信息
                                err.Raise vbObjectError + 101, 100, "Printer Status = " & CStr(printer.printer.Attributes(&H2110, &H10).Value) _
                                    & " Printer Status Info = " & CStr(printer.printer.Attributes(&H2110, &H20).Value)
                            Else
                                'Printer Status出现警告或者错误，但是没有详细信息，直接返回错误或者警告
                                err.Raise vbObjectError + 101, 100, "Printer Status = " & CStr(printer.printer.Attributes(&H2110, &H10).Value)
                            End If
                              
                        End If
                    End If
                End If
            End If
            
            '显示指定页的图像
            Call subShowOnePage(j)
        
            '''''''''''''''''''FilmBox的参数''''''''''''''''''''''''''''''''''
            '胶片方向，必须
            If clsOnePrinter.strOrientation <> "" Then
                printer.Orientation = clsOnePrinter.strOrientation
            Else
                printer.Orientation = "PORTRAIT"
            End If
            '胶片大小，必须
            If clsOnePrinter.strFilmSize <> "" Then
                printer.FilmSize = clsOnePrinter.strFilmSize
            Else
                printer.FilmSize = "14INX17IN"
            End If
            
            '打印图像的位数，必须
            If clsOnePrinter.lngBitDepth <> 0 Then
                printer.BitDepth = clsOnePrinter.lngBitDepth
            Else
                printer.BitDepth = 8
            End If
            
            '放大方式,必须
            If clsOnePrinter.strMagnification <> "" Then
                printer.FilmBox.Attributes.Add &H2010, &H60, clsOnePrinter.strMagnification
            Else
                printer.FilmBox.Attributes.Add &H2010, &H60, "CUBIC"
            End If
            'Smoothing Type '平滑,可选
            If clsOnePrinter.strSmooth <> "" Then
                printer.FilmBox.Attributes.Add &H2010, &H80, clsOnePrinter.strSmooth
            End If
            'border density 边缘密度，必须
            If clsOnePrinter.strBorderDensity <> "" Then
                printer.FilmBox.Attributes.Add &H2010, &H100, clsOnePrinter.strBorderDensity
            Else
                printer.FilmBox.Attributes.Add &H2010, &H100, "BLACK"    'border density
            End If
            'empty image density 空白密度，必须
            If clsOnePrinter.strEmptyDensity <> "" Then
                printer.FilmBox.Attributes.Add &H2010, &H110, clsOnePrinter.strEmptyDensity
            Else
                printer.FilmBox.Attributes.Add &H2010, &H110, "BLACK"   'empty image density
            End If
            'min density 最小密度，必须
        '    If clsOnePrinter.strMinDensity <> "" Then
        '        printer.FilmBox.Attributes.Add &H2010, &H120, clsOnePrinter.strMinDensity
        '    Else
        '        printer.FilmBox.Attributes.Add &H2010, &H120, 16
        '    End If
            'max density 最大密度,必须
        '    If clsOnePrinter.strMaxDensity <> "" Then
        '        printer.FilmBox.Attributes.Add &H2010, &H130, clsOnePrinter.strMaxDensity
        '    Else
        '        printer.FilmBox.Attributes.Add &H2010, &H130, 320
        '    End If
            'trim whether the film will be cut in to 2 or more films 修整胶片,必须
            If clsOnePrinter.strTrim <> "" Then
                printer.FilmBox.Attributes.Add &H2010, &H140, clsOnePrinter.strTrim
            Else
                printer.FilmBox.Attributes.Add &H2010, &H140, "NO"           'trim whether the film will be cut in to 2 or more films
            End If
            'Polarity 极性,可选
            If clsOnePrinter.strPolarity <> "" Then
                printer.FilmBox.Attributes.Add &H2020, &H20, clsOnePrinter.strPolarity
            End If
            'Requested Resolution ID 分辨率，可选
            If clsOnePrinter.strResolution <> "" Then
                printer.FilmBox.Attributes.Add &H2020, &H50, clsOnePrinter.strResolution
            End If
        
            '打印格式，必须
            If marrPages(j).strPageFormat <> "" Then
                printer.Format = marrPages(j).strPageFormat
            Else
                printer.Format = "STANDARD\1,2"
            End If
            
            '计算每一张图的最大分辨率
            Call subCalImageMaxSize(printer.FilmSize, printer.Format, clsOnePrinter.intImageResolution, arrImageSize)
            
            For k = 1 To (FilmViewer.Count - 1)
                If FilmViewer(k).Images.Count > 0 Then
                    Set Thisim = funAssembleImage(FilmViewer(k), arrImageSize(k).intWidth, arrImageSize(k).intHeight)
                    If Not Thisim Is Nothing Then
                        printer.PrintImage Thisim, False, True
                        For i = 1 To FilmViewer(k).Images.Count
                            If InStr(1, StrPrintPage, "," & FilmViewer(k).Images(i).SeriesUID & ",") <= 0 Then
                                StrPrintPage = StrPrintPage & FilmViewer(k).Images(i).SeriesUID & ","
                            End If
                            If InStr(1, strImageUIDS, "," & FilmViewer(k).Images(i).InstanceUID & ",") <= 0 Then
                                strImageUIDS = strImageUIDS & FilmViewer(k).Images(i).InstanceUID & ","
                            End If
                        Next
                    End If
                End If
            Next
            printer.PrintFilm
            printer.Close
            StrPrintLog = StrPrintLog & "|" & Mid(StrPrintPage, 2, Len(StrPrintPage) - 1)
        End If
    Next j
    
    StrPrintLog = Mid(StrPrintLog, 2) '记录胶片使用情况，并标记检查UID已打印
    StrPrintPage = Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, , True).Text
    For i = 0 To UBound(Split(StrPrintLog, "|"))
        strSQL = "Zl_胶片打印记录_Insert('" & Split(StrPrintLog, "|")(i) & "','" & StrPrintPage & "')"
        zlDatabase.ExecuteProcedure strSQL, "记录胶片使用"
    Next
    
    '记录图像打印标记，2000个字符就保存一次，避免一次打印图像太多，超过oracle4000的限制
    If Len(strImageUIDS) > 2000 Then
        arrTempUIDs = Split(strImageUIDS, ",")
        strTempUIDs = ","
        For i = 1 To UBound(arrTempUIDs)
            strTempUIDs = strTempUIDs & arrTempUIDs(i) & ","
            If Len(strTempUIDs) > 2000 Then
                strSQL = "Zl_影像图像胶片打印_Update('" & strTempUIDs & "',1)"
                zlDatabase.ExecuteProcedure strSQL, "记录图像胶片打印情况"
                strTempUIDs = ","
            End If
        Next i
    Else
        strTempUIDs = strImageUIDS
    End If
    If Len(strTempUIDs) > 1 Then
        strSQL = "Zl_影像图像胶片打印_Update('" & strTempUIDs & "',1)"
        zlDatabase.ExecuteProcedure strSQL, "记录图像胶片打印情况"
    End If
    
    '触发打印完成事件
    RaiseEvent AfterPrinted(strImageUIDS)
    
    subPrintFilm = True
    Exit Function
err1:
    MsgBox "打印机连接错误,请检查打印机和网络设置." & vbCrLf & "打印机名为：" & clsOnePrinter.strname & " IP为:" _
            & clsOnePrinter.strIPAddress & " 端口为：" & clsOnePrinter.lngPort & " 错误代码： " & err.Number _
            & " 错误描述： " & err.Description, vbExclamation, gstrSysName, Me
    Exit Function
err2:
    If err.Number = vbObjectError + 101 Then
        MsgBox "打印机没有处于正常状态，返回错误： " & err.Description & ", 请检查打印机后重新打印。", vbExclamation, gstrSysName, Me
    Else
        MsgBox "打印图像过程出现错误，请检查打印格式、胶片大小等设置是否正确。错误代码 ： " & err.Number _
        & " 错误描述： " & err.Description, vbExclamation, gstrSysName, Me
    End If
    On Error Resume Next
    printer.Close
End Function

Public Function funFillPrinterParams(bShowFilmConfig As Boolean) As clsDicomPrint
'------------------------------------------------
'功能：填充一个用于打印的clsDicomPrint打印机类
'参数：bShowFilmConfig－－True则使用胶片设置中的信息填充打印机类；False从打印机列表中查找信息填充打印机类。
'返回：DICOM打印机类
'上级函数或过程：frmFilmConf.cndOK_Click；frmFilm.CommBar_Film_Execute
'下级函数或过程：无
'引用的外部参数：cDICOMPrinter，菜单控件
'编制人：黄捷
'------------------------------------------------
    Dim clsOnePrinter As New clsDicomPrint
    Dim strPrinterName As String
    Dim i As Integer
    
    strPrinterName = Me.CommBar_Film.FindControl(, ID_frmFilm_Camera, , True).Text
    '判断打印机是否存在
    For i = 1 To cDICOMPrinter.Count
        If cDICOMPrinter(i).strname = strPrinterName Then
            Exit For
        End If
    Next i
    
    If i > cDICOMPrinter.Count Then
        MsgBox "打印机：" & strPrinterName & " 没有找到。", vbInformation, gstrSysName, Me
        Exit Function
    End If
    '将打印机的设置保存到clsOnePrinter中
    
    Set clsOnePrinter = cDICOMPrinter(strPrinterName)
    If bShowFilmConfig Then              '从胶片设置中获取信息
        With clsOnePrinter
            .strFilmBox = frmFilmConf.cboFilmBox.Text
            .strFilmSize = frmFilmConf.cboFilmSize.Text
            .strFormat = frmFilmConf.cboFormat.Text
            .strMagnification = frmFilmConf.cboMagnification.Text
            .strMedium = frmFilmConf.cboMedium.Text
            .strOrientation = frmFilmConf.cboOrientation.Text
            .strPriority = frmFilmConf.cboPriority.Text
            .strResolution = frmFilmConf.cboResolution.Text
            .strSmooth = frmFilmConf.cboSmooth.Text
            .strTrim = frmFilmConf.cboTrim.Text
            .lngCopies = frmFilmConf.lstCopies.list(frmFilmConf.lstCopies.TopIndex)
        End With
        mintPageRange = frmFilmConf.cboPageRange.ListIndex
    Else
        With clsOnePrinter
            .strOrientation = IIf(mblnIsPortrait, "PORTRAIT", "LANDSCAPE")
            .strFilmSize = CommBar_Film.FindControl(, ID_frmFilm_FilmSize, True).Text
            .strFormat = CommBar_Film.FindControl(, ID_frmFilm_Format, True).Text
        End With
        mintPageRange = 0
    End If
    Set funFillPrinterParams = clsOnePrinter
    
End Function

Private Sub CreateBar()
    '------------------------------------------------
    '功能：                                  创建菜单
    '参数：
    '返回：                                  无
    '------------------------------------------------
    Dim ToolBar As CommandBar
    Dim control As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cboControl As CommandBarComboBox
    
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.CommBar_Film.VisualTheme = xtpThemeOffice2003
    Me.CommBar_Film.Icons = ImgIcons.Icons
    
    With Me.CommBar_Film.Options
        .ShowExpandButtonAlways = False     '去掉扩展按钮
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 32, 32
    End With
    
    'Me.CommBar_Film.VisualTheme = IntComBarTheme                            '统一工具条风格
    Me.CommBar_Film.Item(1).Visible = False                                 '隐藏菜单栏
    
    '建立主工具栏
    Set ToolBar = Me.CommBar_Film.Add("主工具栏", mintTBMainPosition)
    Call ToolBar.EnableDocking(xtpFlagAlignTop)
    
    With ToolBar.Controls
        Set control = .Add(xtpControlButton, ID_frmFilm_TakePictures, "照相")
        control.Style = xtpButtonIconAndCaption ' xtpButtonIcon 'cbrControl.style = xtpButtonIconAndCaption
        control.IconId = 1001
        
        Set control = .Add(xtpControlButton, ID_frmFilm_FilmCol, "纵向")
        control.BeginGroup = True
        .Add xtpControlButton, ID_frmFilm_FilmRow, "横向"
        .Add xtpControlButton, ID_frmFilm_RectPhotCase, "正方格图像"
        .Add xtpControlButton, ID_frmFilm_FormatCustom, "格式定义"
        
        Set control = .Add(xtpControlButton, ID_frmFilm_OpenImages, "打开")
        control.Style = xtpButtonIconAndCaption
        control.IconId = 1005
        control.ToolTipText = "打开图像"
        control.BeginGroup = True
        Set control = .Add(xtpControlButton, ID_frmFilm_DeleteImg, "删除")
        control.Style = xtpButtonIconAndCaption
        control.IconId = 1002
        control.ToolTipText = "删除图象"
        
        Set control = .Add(xtpControlComboBox, ID_frmFilm_FilmSize, "胶片大小")
        control.BeginGroup = True
        Set cboControl = .Add(xtpControlComboBox, ID_frmFilm_Format, "格式")
        cboControl.width = 120
        Set control = .Add(xtpControlButton, ID_frmFilm_Divide, "图像分格")
        control.Style = xtpButtonIconAndCaption
        control.IconId = 1004
        
        .Add xtpControlComboBox, ID_frmFilm_Camera, "相机"
        Set control = .Add(xtpControlButton, ID_frmFilm_Quit, "退出")
        control.Style = xtpButtonIconAndCaption
        control.IconId = 1003
    End With
        
    '建立图像操作工具栏
    Set ToolBar = Me.CommBar_Film.Add("图像操作栏", mintTBImageProcessPosition)
    Call ToolBar.EnableDocking(xtpFlagAlignAny)
    ToolBar.ShowTextBelowIcons = True
    
    With ToolBar.Controls
        
        Set control = .Add(xtpControlSplitButtonPopup, ID_Active_AdjustWindow_HandAdjustWindow, " 调窗")
        
        control.IconId = 1006
        control.BeginGroup = True
        
        Set control = .Add(xtpControlButton, ID_frmFilm_Pan, " 漫游")
        control.IconId = 1008
        Set control = .Add(xtpControlButton, ID_frmFilm_Zoom, " 缩放")
        control.IconId = 1007
        Set control = .Add(xtpControlButton, ID_frmFilm_Invert, " 反白")
        control.IconId = 1009
        Set control = .Add(xtpControlButton, ID_frmFilm_RotateLeft, " 左旋")
        control.IconId = 1010
        Set control = .Add(xtpControlButton, ID_frmFilm_RotateRight, " 右旋")
        control.IconId = 1011
        Set control = .Add(xtpControlButton, ID_frmFilm_FilterLengthDown, " 平滑减")
        control.IconId = 1012
        control.BeginGroup = True
        Set control = .Add(xtpControlButton, ID_frmFilm_FilterLengthUp, " 平滑加")
        control.IconId = 1013
        Set control = .Add(xtpControlButton, ID_frmFilm_RectZoom, " 框选")
        control.IconId = 1014
        control.BeginGroup = True
        
        Set control = .Add(xtpControlSplitButtonPopup, ID_frmFilm_CutOut, " 裁剪")
        control.BeginGroup = True
        control.ToolTipText = "自由比例裁剪，拖动鼠标左键，选择裁剪区域，双击图像进行裁剪"
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_Custom, "自由裁剪")
        cbrPopControl.Checked = True
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_14X17, "14*17")
        cbrPopControl.BeginGroup = True
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_11X14, "11*14")
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_10X14, "10*14")
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_8X10, "8*10")
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_14X14, "14*14")
        cbrPopControl.BeginGroup = True
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_17X14, "17*14")
        cbrPopControl.BeginGroup = True
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_14X11, "14*11")
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_14X10, "14*10")
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_10X8, "10*8")
        
        Set control = .Add(xtpControlButton, ID_frmFilm_FlipHorizontal, "镜象")
        control.IconId = 1016
        Set control = .Add(xtpControlButton, ID_frmFilm_FlipVertical, "倒置")
        control.IconId = 1017
        Set control = .Add(xtpControlButtonPopup, ID_frmFilm_Label, "标注")
        control.IconId = 1018
        control.BeginGroup = True
        control.Id = ID_frmFilm_Label
        control.CommandBar.Controls.Add xtpControlButton, ID_frmFilm_Label_L, "L(左)"
        control.CommandBar.Controls.Add xtpControlButton, ID_frmFilm_Label_R, "R(右)"
        control.CommandBar.Controls.Add xtpControlButton, ID_frmFilm_Label_A, "A(前)"
        control.CommandBar.Controls.Add xtpControlButton, ID_frmFilm_Label_P, "P(后)"
        control.CommandBar.Controls.Add xtpControlButton, ID_frmFilm_Label_S, "S(上)"
        control.CommandBar.Controls.Add xtpControlButton, ID_frmFilm_Label_I, "I(下)"
        control.CommandBar.Controls.Add xtpControlButton, ID_frmFilm_Label_Delete, "清除标注"
        
        Set control = .Add(xtpControlButton, ID_frmFilm_SelAll, "全选")
        control.IconId = 1020
        control.BeginGroup = True
        Set control = .Add(xtpControlButton, ID_frmFilm_SelSeries, "序列选")
        control.IconId = 1021
        Set control = .Add(xtpControlButton, ID_frmFilm_SelInverse, "反选")
        control.IconId = 1022
        Set control = .Add(xtpControlButton, ID_frmFilm_SelNone, "全清")
        control.IconId = 1023
        
        Set control = .Add(xtpControlButton, ID_frmFilm_ImgIncrease, "正序")
        control.BeginGroup = True
        Set control = .Add(xtpControlButton, ID_frmFilm_ImgDecrease, "逆序")
        
        Set control = .Add(xtpControlButton, ID_frmFilm_Resume, " 恢复 ")
        control.IconId = 1019
        
        control.BeginGroup = True
        
        .Add xtpControlButton, ID_frmFilm_UnDivide, "合并测试"
    End With
    
    Me.CommBar_Film.FindControl(, ID_frmFilm_FilmCol, , True).Checked = True
    Me.CommBar_Film.FindControl(, ID_frmFilm_RectPhotCase, , True).Checked = False
End Sub

Public Sub subSynchronalImg(blnRestore As Boolean, intType As Integer)
'------------------------------------------------
'功能：对胶片打印中的图像处理做同步
'参数： blnRestore -True 恢复原图像参数；False - 按照选定的图像同步
'       intType   --- 图像同步类型，宏定义
'返回：无
'------------------------------------------------
    
    If (Not SelectedImage Is Nothing) And funGetTagVal(SelectedImage.Tag, TAG_选择) = "Select" Then
        Dim v As DicomViewer
        Dim img As DicomImage
        Dim i As Integer
        
        For Each v In FilmViewer
            For i = 1 To v.Images.Count
                If funGetTagVal(v.Images(i).Tag, TAG_选择) = "Select" Then
                        Set img = v.Images(i)
                        If blnRestore = True Then
                            img.SetDefaultWindows
                            img.StretchToFit = True
                            img.FlipState = doFlipNormal
                            img.RotateState = doRotateNormal
                            img.UnsharpEnhancement = 0
                            img.UnsharpLength = 0
                            img.FilterLength = 0
                            If img.Labels.Count >= G_INT_SYS_LABEL_WWWL Then
                                img.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & img.width & "-L:" & img.Level
                            End If
                            If img.Labels.Count >= G_INT_SYS_LABEL_RULLER Then
                                If img.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then  '更新标尺单位
                                    UpdateRuler img, True
                                End If
                            End If
                        Else
                            Call subImageInPhase(img, SelectedImage, intType)
                        End If
                End If
            Next i
        Next v
    End If
    '把修改过的图像上传到原始图集合中
    Call subReloadImgsPrint
End Sub

Private Sub subReloadImgsPrint()
'------------------------------------------------
'功能：将修改过显示参数的图像重新加载回imgsPrint图像集中
'参数：无
'返回：无
'上级函数或过程：
'下级函数或过程：
'引用的外部参数：
'编制人：黄捷 2006-2-17
'------------------------------------------------
    Dim v As DicomViewer
    Dim img As DicomImage
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim intStart As Integer
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    If FilmViewer.Count <= 1 Then Exit Sub
    '计算开始装载的第一个图像编号
    intStart = funGetStartImgNo(VScro.Value, 1, 1)
    For i = 1 To FilmViewer.Count - 1
        lngWidth = FilmViewer(i).width / FilmViewer(i).MultiColumns
        lngHeight = FilmViewer(i).height / FilmViewer(i).MultiRows
        
        For j = 1 To FilmViewer(i).Images.Count
            '删除imgsPrint中对应的图像
            imgsPrint.Remove intStart
            '往imgsPrint中增加图像
            '记录图像所在viewer所占用的原始宽度和高度
            Call funSetTagVal(FilmViewer(i).Images(j), TAG_V宽度, CStr(lngWidth))
            Call funSetTagVal(FilmViewer(i).Images(j), TAG_V高度, CStr(lngHeight))
            imgsPrint.Add FilmViewer(i).Images(j)
            '将新增到imgsPrint中的图像移动到原有的位置
            imgsPrint.Move imgsPrint.Count, intStart
            intStart = intStart + 1
        Next j
    Next i
End Sub

Private Function funAssembleImage(AssembleViewer As DicomViewer, Optional intImgMaxWidth As Integer = 0, _
    Optional intImgMaxHeight As Integer = 0) As DicomImage
'------------------------------------------------
'功能：组合viewer中的显示的所有图像成一个图像
'参数： AssembleViewer--需要组合的Viewer
'       intImgMaxWidth -- 组合图像的最大宽度
'       intImgMaxHeight -- 组合图像的最大高度
'返回：返回组合好的图像
'------------------------------------------------
    Dim Image As New DicomImage '新图像
    Dim imgs As New DicomImages '临时存储屏幕采集的图像集
    Dim intWidth As Integer     '新图像的宽度
    Dim intHeight As Integer    '新图像的高度
    Dim Simg As New DicomImage
    Dim intLeft As Integer
    Dim intRight As Integer
    Dim intTop As Integer
    Dim intBottom As Integer
    Dim sZoom As Single
    Dim sOldZoom As Single          '记录原来的缩放倍数
    Dim intImgRectWidth As Integer  '单张图像可占用的区域宽度
    Dim intImgRectHeight As Integer '单张图像可占用的区域高度
    Dim i As Integer
    Dim intMaxWidth As Integer      '拼接后图像的最大宽度
    Dim intMaxHeight As Integer     '拼接后图像的最大高度
    Dim intBorder As Integer        '图像之间的边距
    Dim intImgX As Integer          'X方向的图像数量
    Dim intImgY As Integer          'Y方向的图像数量
    Dim intActualSizex As Integer   '图像旋转变换后X方向的像素点数
    Dim intActualSizey As Integer   '图像旋转变换后Y方向的像素点数
    Dim intOffsetX As Integer       '拼接时X方向的位移
    Dim intOffsetY As Integer       '拼接时Y方向的位移
    Dim dlImgLabel As DicomLabel    '图像的标注
    Dim j As Integer
    Dim dblX As Double, dblY As Double, intTemp As Integer
    Dim iMaxWidth As Integer, iMaxHeight As Integer
    Dim dblScaleZoom As Double
    Dim lngTempHeight  As Long
    Dim lngTempWidth As Long
    Dim lngImgLeft As Long
    Dim lngImgTop As Long
    Dim strPatiInfo(4) As String
        
    On Error GoTo err
    
    If AssembleViewer.Images.Count <= 0 Then
        '返回一个黑图**************
        Exit Function
    End If
        
    
    '计算新图像的宽度和高度
    '新图像的宽度和高度不能够大于intMaxWidth×intMaxHeight（宽度×高度）
    If intImgMaxWidth = 0 Then
        intMaxWidth = 3073
    Else
        intMaxWidth = intImgMaxWidth
    End If
    
    If intImgMaxHeight = 0 Then
        intMaxHeight = 3073
    Else
        intMaxHeight = intImgMaxHeight
    End If
    
    intBorder = 10
    intImgRectWidth = 0
    intImgRectHeight = 0
    
    '估算新图像的宽度和高度
    '使用原图像的宽度和高度和，并用Viewer的比例来修正。
    '估算图像的新宽高
    For i = 1 To AssembleViewer.Images.Count
        '查找旋转变换后图像的x方向点数
        intActualSizex = AssembleViewer.Images(i).sizex
        If AssembleViewer.Images(i).RotateState = doRotateLeft Or AssembleViewer.Images(i).RotateState = doRotateRight Then
            intActualSizex = AssembleViewer.Images(i).sizey
        End If
        
        '查找旋转变换后图像的y方向点数
        intActualSizey = AssembleViewer.Images(i).sizey
        If AssembleViewer.Images(i).RotateState = doRotateLeft Or AssembleViewer.Images(i).RotateState = doRotateRight Then
            intActualSizey = AssembleViewer.Images(i).sizex
        End If
        
        If intImgRectWidth < intActualSizex Then intImgRectWidth = intActualSizex
        If intImgRectHeight < intActualSizey Then intImgRectHeight = intActualSizey
    Next i
    
    '计算横向和纵向图像数量
    intImgX = AssembleViewer.Images.Count
    If intImgX > AssembleViewer.MultiColumns Then intImgX = AssembleViewer.MultiColumns
    intImgY = (AssembleViewer.Images.Count - 1) \ AssembleViewer.MultiColumns + 1
    
    '修正单个图像的最大区域
    If intImgRectWidth > intMaxWidth / intImgX Or intImgRectHeight > intMaxHeight / intImgY Then
        intImgRectWidth = intMaxWidth / intImgX
        intImgRectHeight = intMaxHeight / intImgY
    End If
    
    intWidth = intImgRectWidth * intImgX
    intHeight = intImgRectHeight * intImgY
    
    '修正图像的宽高，不能大于最大值
    '如果大于intMaxWidth×intMaxHeight则，按照图像总长宽比，使用小于等于intMaxWidth×intMaxHeight作为新宽高,
    If intWidth > intMaxWidth Or intHeight > intMaxHeight Then
        If intHeight / intWidth > intMaxHeight / intMaxWidth Then
            intWidth = intWidth / intHeight * intMaxHeight
            intHeight = intMaxHeight
        Else
            intHeight = intHeight / intWidth * intMaxWidth
            intWidth = intMaxWidth
        End If
    End If
    
    '采集图像
    '将图像采集到临时图像集
    lngImgTop = 1
    lngImgLeft = 1
    For i = 1 To AssembleViewer.Images.Count
        '计算采集图像的大小
        
        intLeft = AssembleViewer.ImageXPosition(lngImgLeft, lngImgTop)
        intTop = AssembleViewer.ImageYPosition(lngImgLeft, lngImgTop)
        
         '计算下一个图的Left，Top,计算位置时，有小数就进位（+0.5），防止图像多了，累计位移出现偏差。
        lngImgLeft = lngImgLeft + AssembleViewer.width / AssembleViewer.MultiColumns / Screen.TwipsPerPixelX + 0.5
        If lngImgLeft >= AssembleViewer.width / Screen.TwipsPerPixelX Then
            lngImgLeft = 1
            lngImgTop = lngImgTop + AssembleViewer.height / AssembleViewer.MultiRows / Screen.TwipsPerPixelY + 0.5
        End If
        
        If AssembleViewer.Images(i).RotateState = doRotateLeft Or AssembleViewer.Images(i).RotateState = doRotateRight Then
            lngTempWidth = AssembleViewer.height / AssembleViewer.MultiRows / Screen.TwipsPerPixelY / AssembleViewer.Images(i).ActualZoom
            lngTempHeight = AssembleViewer.width / AssembleViewer.MultiColumns / Screen.TwipsPerPixelX / AssembleViewer.Images(i).ActualZoom
        Else
            lngTempHeight = AssembleViewer.height / AssembleViewer.MultiRows / Screen.TwipsPerPixelY / AssembleViewer.Images(i).ActualZoom
            lngTempWidth = AssembleViewer.width / AssembleViewer.MultiColumns / Screen.TwipsPerPixelX / AssembleViewer.Images(i).ActualZoom
        End If
        
        If (AssembleViewer.Images(i).RotateState = doRotateLeft And (AssembleViewer.Images(i).FlipState = doFlipNormal Or AssembleViewer.Images(i).FlipState = doFlipVertical)) Or _
            (AssembleViewer.Images(i).RotateState = doRotateRight And (AssembleViewer.Images(i).FlipState = doFlipBoth Or AssembleViewer.Images(i).FlipState = doFlipHorizontal)) Or _
            (AssembleViewer.Images(i).RotateState = doRotate180 And (AssembleViewer.Images(i).FlipState = doFlipVertical Or AssembleViewer.Images(i).FlipState = doFlipNormal)) Or _
            (AssembleViewer.Images(i).RotateState = doRotateNormal And (AssembleViewer.Images(i).FlipState = doFlipHorizontal Or AssembleViewer.Images(i).FlipState = doFlipBoth)) Then
            intLeft = intLeft - lngTempWidth
        End If
        
        If (AssembleViewer.Images(i).RotateState = doRotateLeft And (AssembleViewer.Images(i).FlipState = doFlipVertical Or AssembleViewer.Images(i).FlipState = doFlipBoth)) Or _
            (AssembleViewer.Images(i).RotateState = doRotateRight And (AssembleViewer.Images(i).FlipState = doFlipNormal Or AssembleViewer.Images(i).FlipState = doFlipHorizontal)) Or _
            (AssembleViewer.Images(i).RotateState = doRotate180 And (AssembleViewer.Images(i).FlipState = doFlipNormal Or AssembleViewer.Images(i).FlipState = doFlipHorizontal)) Or _
            (AssembleViewer.Images(i).RotateState = doRotateNormal And (AssembleViewer.Images(i).FlipState = doFlipVertical Or AssembleViewer.Images(i).FlipState = doFlipBoth)) Then
            intTop = intTop - lngTempHeight
        End If
        
        intRight = lngTempWidth + intLeft
        intBottom = lngTempHeight + intTop

        '查找旋转变换后图像的x方向点数
        intActualSizex = AssembleViewer.Images(i).sizex
        If AssembleViewer.Images(i).RotateState = doRotateLeft Or AssembleViewer.Images(i).RotateState = doRotateRight Then
            intActualSizex = AssembleViewer.Images(i).sizey
        End If
        
        '查找旋转变换后图像的y方向点数
        intActualSizey = AssembleViewer.Images(i).sizey
        If AssembleViewer.Images(i).RotateState = doRotateLeft Or AssembleViewer.Images(i).RotateState = doRotateRight Then
            intActualSizey = AssembleViewer.Images(i).sizex
        End If
        
        '计算缩放比例 hj修改,解决多图合并时，放大的图象无法真正放大的问题
        sZoom = intImgRectHeight / IIf((intBottom - intTop) > intActualSizey Or intBottom = 0, intActualSizey, (intBottom - intTop))
        If sZoom > intImgRectWidth / IIf((intRight - intLeft) > intActualSizex Or intRight = 0, intActualSizex, (intRight - intLeft)) Then
            sZoom = intImgRectWidth / IIf((intRight - intLeft) > intActualSizex Or intRight = 0, intActualSizex, (intRight - intLeft))
        End If
      
        '把图像按照新比例缩放，然后重新划标尺
        sOldZoom = AssembleViewer.Images(i).ActualZoom
        AssembleViewer.Images(i).StretchToFit = False
        AssembleViewer.Images(i).Zoom = sZoom
        
        If UpdateRuler(AssembleViewer.Images(i), True) = 1 Then
            '标尺标注不存在，可能是“合并测试”之后的图像，再次进行“合并测试”。
            MsgBox "图像的标尺信息不存在。" & vbCrLf & vbCrLf & "原因分析：" & vbCrLf & "    1、合并测试的结果图像，不能再次进行合并测试。" & vbCrLf & "    2、其他未知错误。", vbOKOnly, "错误提示"
            Exit Function
        End If
        
        '先设置打印字体大小，设置用户自己画的标注
        Call subChangeLabelForPrint(AssembleViewer.Images(i), 1)

        '隐藏图像的四角信息
        Call subDispImageInfo(AssembleViewer.Images(i), False, False, True)
                
        Set Simg = AssembleViewer.Images(i).PrinterImage(8, 1, True, sZoom, intLeft, intRight, intTop, intBottom)
        
        
        '显示图像的四角信息
        '把原来图像的标注，添加到现在的图像中，因为用户自己画的标注在上一步已经被画到图像中，因此这里只恢复系统标注
        Simg.Labels.Clear
        For j = 1 To IIf(G_INT_SYS_LABEL_COUNT <= AssembleViewer.Images(i).Labels.Count, G_INT_SYS_LABEL_COUNT, AssembleViewer.Images(i).Labels.Count)
            Simg.Labels.Add AssembleViewer.Images(i).Labels(j)
            Simg.Labels(Simg.Labels.Count).Visible = False
        Next j
        '添加原来的图像类型
        Simg.Attributes.Add &H8, &H60, AssembleViewer.Images(i).Attributes(&H8, &H60)
        Call subDispImageInfo(Simg, True, False, False)     ''显示病人四角信息和窗宽窗位信息
        '因为标尺已经在前面画过了，因此这里隐藏标尺的显示
        Call UpdateRuler(Simg, False)
        
        '设置打印字体大小
        Call subChangeLabelForPrint(Simg, 1)

        '计算图像显示区域的比例，把图像摆在中间，四角文字放到可视区域的四个角
        dblX = Simg.sizex / (AssembleViewer.width / AssembleViewer.MultiColumns)
        dblY = Simg.sizey / (AssembleViewer.height / AssembleViewer.MultiRows)
        If dblX < dblY Then
            intTemp = dblY * AssembleViewer.width / AssembleViewer.MultiColumns
            intLeft = -(intTemp - Simg.sizex) / 2
            intRight = intTemp + intLeft
            intTop = 0
            intBottom = 0
        Else
            intTemp = dblX * AssembleViewer.height / AssembleViewer.MultiRows
            intLeft = 0
            intRight = 0
            intTop = -(intTemp - Simg.sizey) / 2
            intBottom = intTemp + intTop
        End If
        
        Set Simg = Simg.PrinterImage(8, 1, True, 1, intLeft, intRight, intTop, intBottom)
        
        '恢复图像原来的缩放比例
        AssembleViewer.Images(i).Zoom = sOldZoom
        
        '恢复图像原来的标注
        '显示图像的四角信息
        Call subDispImageInfo(AssembleViewer.Images(i), True, False, True)
        
        '设置打印字体大小，设置用户自己画的标注
        Call subChangeLabelForPrint(AssembleViewer.Images(i), 0)

        imgs.Add Simg
    Next i
     
    '精确计算新图像的宽度和高度
    intImgRectWidth = 0
    intImgRectHeight = 0
     
    For i = 1 To imgs.Count
        If intImgRectWidth < imgs(i).sizex Then intImgRectWidth = imgs(i).sizex
        If intImgRectHeight < imgs(i).sizey Then intImgRectHeight = imgs(i).sizey
    Next i
    
    If Not clsTruePrinter Is Nothing Then
        intBorder = clsTruePrinter.lngImageBorderWidth
    Else
        intBorder = 1
    End If
     
    intImgRectWidth = intImgRectWidth + intBorder
    intImgRectHeight = intImgRectHeight + intBorder
    intWidth = intImgRectWidth * intImgX
    intHeight = intImgRectHeight * intImgY
    
    '创建新图像
    Image.Name = "print"
    Image.PatientID = "print001"
    Image.Attributes.Add &H8, &H16, doSOP_SecondaryCapture
    Image.Attributes.Add &H28, &H2, 1 ' samples/pixel
    Image.Attributes.Add &H28, &H4, "MONOCHROME2" ' photometric interpreation  'CT都是MONOCHROME2,CR都是MONOCHROME1？
    Image.Attributes.Add &H28, &H10, intHeight   'x,Rows
    Image.Attributes.Add &H28, &H11, intWidth 'Y,Columns
    Image.Attributes.Add &H28, &H100, 8 'bits allocated
    Image.Attributes.Add &H28, &H101, 8 ' bits stored
    Image.Attributes.Add &H28, &H102, 7 ' high bit
    ReDim pix(intWidth, intHeight) As Byte
    
    Image.Attributes.Add &H7FE0, &H10, pix
    
    '获取最大的图像宽度和高度
    iMaxWidth = 0
    iMaxHeight = 0
    For i = 1 To imgs.Count
        If iMaxWidth < imgs(i).sizex Then
            iMaxWidth = imgs(i).sizex
            iMaxHeight = imgs(i).sizey
        End If
    Next i
    
    '拼接新图像
    For i = 1 To imgs.Count
        '计算图像内位移
        dblScaleZoom = iMaxWidth / imgs(i).sizex
        intOffsetX = (intImgRectWidth - imgs(i).sizex * dblScaleZoom) / 2
        intOffsetY = (intImgRectHeight - imgs(i).sizey * dblScaleZoom) / 2
        Image.Blt imgs(i), 0, 0, ((i - 1) Mod AssembleViewer.MultiColumns) * intImgRectWidth + intOffsetX, ((i - 1) \ AssembleViewer.MultiColumns) * intImgRectHeight + intOffsetY, imgs(i).sizex * dblScaleZoom, imgs(i).sizey * dblScaleZoom, 1, 1, dblScaleZoom, False
    Next i
    
    Set funAssembleImage = Image
    Exit Function
    
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function funGetStartViewerNo(intPage As Integer, intViewer As Integer) As Integer
'------------------------------------------------
'功能：通过当前页数，获取从第一页到当前中，Viewer的总数。
'参数： intPage     --当前页数索引
'       intViewer   --当前Viewer索引
'返回：在intViewer 之前的所有viewer的总数
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
    
    funGetStartViewerNo = 0
    
    If intPage > UBound(marrPages) Then Exit Function
    If intViewer > marrPages(intPage).intViewerCount Then Exit Function
    
    For i = 1 To intPage - 1
        funGetStartViewerNo = funGetStartViewerNo + marrPages(i).intViewerCount
    Next i
    
    funGetStartViewerNo = funGetStartViewerNo + intViewer
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    funGetStartViewerNo = 0
End Function

Private Function funGetStartImgNo(intPage As Integer, intViewer As Integer, intImage As Integer) As Integer
'------------------------------------------------
'功能：通过当前页数，获取从第一页到当前页之前，Images的总数。
'参数： intPage     --当前页数索引
'       intViewer   --当前Viewer的索引
'       intImage    --当前Image的索引
'返回：当前页intImage之前的Images的总数
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
    
    funGetStartImgNo = 0
    
    If intPage > UBound(marrPages) Then Exit Function
    If intViewer > marrPages(intPage).intViewerCount Then Exit Function
    If intImage > marrPages(intPage).ViewerLayout(intViewer).intColumns * marrPages(intPage).ViewerLayout(intViewer).intRows Then Exit Function
    
    For i = 1 To intPage - 1
        funGetStartImgNo = funGetStartImgNo + marrPages(i).intImageCount
    Next i
    
    For i = 1 To intViewer - 1
        funGetStartImgNo = funGetStartImgNo + marrPages(intPage).ViewerLayout(i).intColumns * marrPages(intPage).ViewerLayout(i).intRows
    Next i
    
    funGetStartImgNo = funGetStartImgNo + intImage
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    funGetStartImgNo = 0
End Function

Public Sub subDispReferLineFilm()
'------------------------------------------------
'功能：在胶片打印窗体中显示定位线
'参数：
'返回值：无
'2009 用
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim v As DicomViewer
    Dim im As DicomImage
    Dim imm As DicomImage
    
    On Error Resume Next
    
    '先删除原来的定位线
    For Each v In Me.FilmViewer
        For Each im In v.Images
            subDeleteAppointLabel im, "RL"          '删除指定类型的标注
        Next
        v.Refresh
    Next
    
    '显示所有定位线
    If Button_miAllReferLine = False Then Exit Sub
    
    For i = 1 To (FilmViewer.Count - 1)
        For Each im In Me.FilmViewer(i).Images
            If subGetReferImg(im) = True Then
                For j = 1 To (FilmViewer.Count - 1)
                    For Each imm In Me.FilmViewer(j).Images
                        If subGetReferImg(imm) = False Then
                            Call subDrawRefLine(imm, im, True, "RLL", False)
                        End If
                    Next
                Next
            End If
        Next
        Me.FilmViewer(i).Refresh
    Next
End Sub
Private Function subGetReferImg(img As DicomImage) As Boolean
    '功能  当前图像是否是定位像
    '参数： img --- 需要判断的图像
    '返回： True -- 是定位像；False -- 不是定位像
    
    Dim v As Variant
    Dim i As Integer
    Dim strAttr As String
    
    On Error GoTo err
    v = img.Attributes(&H8, &H8).Value
    
    If (VarType(v) > 8192) Then
        For i = LBound(v, 1) To UBound(v, 1)
            If i = 3 And v(i) = "LOCALIZER" Then
                subGetReferImg = True
                Exit Function
            End If
        Next
    End If
    
    '如果（8,8）中没有“LOCALIZER”标记，再检查图像的序列描述(8,103E)中是否包含“LOC”
    'GE 的MR，序列描述中包含“LOC”；飞利浦的MR，序列描述中包含"SURVEY"
    If img.Attributes(&H8, &H103E).Exists And Not IsNull(img.Attributes(&H8, &H103E).Value) Then
        strAttr = img.Attributes(&H8, &H103E).Value
        If InStr(UCase(strAttr), "LOC") <> 0 Or InStr(UCase(strAttr), "SURVEY") <> 0 Then
            subGetReferImg = True
            Exit Function
        End If
    End If
    
err:
    '出错不处理
End Function

Private Sub subDeleteAllImages()
'------------------------------------------------
'功能： 删除全部图像
'参数：无
'返回：无
'------------------------------------------------
    
    On Error GoTo err
    
    '清空真实图像
    imgsPrint.Clear
        
    '删除图像之后，图像减少了，重新计算页数
    Call subRecalPages
    '重新显示图像
    Call subShowPrintImages(1)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub CommBar_Execute_PrintFilm()
    Dim strPrinterName As String
    Dim i As Integer
    
    If mblnPrinted = True Then
        If MsgBox("该胶片已经打印过了，是否需要再次打印？", vbYesNo, gstrSysName, Me) = vbNo Then
            Exit Sub
        End If
    End If
        
    strPrinterName = Me.CommBar_Film.FindControl(, ID_frmFilm_Camera, , True).Text
    If Len(Trim(strPrinterName)) <= 0 Then
        MsgBox "请选择胶片打印机!", vbInformation, gstrSysName, Me
        Exit Sub
    End If
    '检查打印机是否存在
    For i = 1 To cDICOMPrinter.Count
        If cDICOMPrinter(i).strname = strPrinterName Then
            Exit For
        End If
    Next i
    
    If i > cDICOMPrinter.Count Then
        MsgBox "打印机：" & strPrinterName & " 没有找到。" & vbCrLf & vbCrLf & "请选择胶片打印机!", vbInformation, gstrSysName, Me
        Exit Sub
    End If
        
    '判断是否需要显示打印设置窗口
    If bShowFilmConfig Then
        Set frmFilmConf.f = Me
        With frmFilmConf
            .sstabFilmConfig.TabVisible(0) = False
            .sstabFilmConfig.TabVisible(1) = True
            If strPrinterName = "" Then Exit Sub
            .cboFilmBox.Text = cDICOMPrinter(strPrinterName).strFilmBox
            .cboFilmSize.Text = Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, , True).Text
            .cboFormat.Text = Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text
            .cboMagnification.Text = cDICOMPrinter(strPrinterName).strMagnification
            .cboMedium.Text = cDICOMPrinter(strPrinterName).strMedium
            .cboOrientation.Text = IIf(mblnIsPortrait, "PORTRAIT", "LANDSCAPE")
            .cboPriority.Text = cDICOMPrinter(strPrinterName).strPriority
            .cboResolution.Text = cDICOMPrinter(strPrinterName).strResolution
            .cboSmooth.Text = cDICOMPrinter(strPrinterName).strSmooth
            .cboTrim.Text = cDICOMPrinter(strPrinterName).strTrim
            .lstCopies = cDICOMPrinter(strPrinterName).lngCopies
            If .zlShowMe = False Then
                Exit Sub
            End If
        End With
    Else
        If MsgBox("是否打印胶片？", vbOKCancel, "PACS提示", Me) = vbCancel Then
            Exit Sub
        End If
        Set clsTruePrinter = funFillPrinterParams(bShowFilmConfig)
    End If
        
    '打印之前，锁定程序窗口
    mblnPrinting = True
    On Error GoTo err
    
    CommBar_Film.FindControl(, ID_frmFilm_TakePictures, , True).Enabled = False
    Me.Caption = "正在打印胶片，请稍候......"
    If subPrintFilm(clsTruePrinter) = True Then
        mintPrintFilmCount = mintPrintFilmCount + 1
        Me.Caption = "胶片打印预览，打印第 " & mintPrintFilmCount & " 次"
        If blnPrintOkEcho = True Then
            MsgBox "胶片打印预览，打印第 " & mintPrintFilmCount & " 次成功。", vbOKOnly, gstrSysName, Me
        End If
        
        '打印成功，则清空图像
        If mblnClearAfterPrint = True Then
            Call subDeleteAllImages
        End If
        
        '提示打印完成的声音
        Call PrintFilmBeep(2)
    Else
        Me.Caption = "胶片打印不成功，请重新设置后再打印"
    End If
    mblnPrinting = False
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    mblnPrinting = False
End Sub

Private Sub OpenImages(strImageIDs As String)
'------------------------------------------------
'功能：打开图像选择窗口，允许用户自己选择图像添加到胶片中
'参数： strImageIDs -- 要打开的序列ID串，'规则是“序列号1|1-3;5-27;33-100+序列号2|全部”,全部表示打开全部图象
'返回：无
'------------------------------------------------
    Dim blnAllImages As Boolean         '是否打开全部图像
    Dim imgs As New DicomImages              '需要打开的图像集
    Dim iSeriesID As Integer            '序列号
    Dim strSeries() As String
    Dim strImages() As String
    Dim i As Integer
    Dim j As Integer
    Dim k  As Integer
    Dim h As Integer
    Dim tmpImg As DicomImage
    
    On Error GoTo err
    
    '把选中的图像添加到imgs 中
    strSeries = Split(strImageIDs, "+")
    For i = 0 To UBound(strSeries)
        iSeriesID = Split(strSeries(i), "|")(0)
        If Split(strSeries(i), "|")(1) = "全部" Then
            For k = 1 To ZLSeriesInfos(iSeriesID).ImageInfos.Count
                Set tmpImg = funLoadAImage(iSeriesID, k, 0)
                
                If Not tmpImg Is Nothing Then
                    subInitAImage tmpImg, 0, Nothing
                    imgs.Add tmpImg
                End If
            Next k
        Else
            strImages = Split(Split(strSeries(i), "|")(1), ";")
            For j = 0 To UBound(strImages)
                For h = Split(strImages(j), "-")(0) To Split(strImages(j), "-")(1)
                    Set tmpImg = funLoadAImage(iSeriesID, h, 0)
                    If Not tmpImg Is Nothing Then
                        subInitAImage tmpImg, 0, Nothing
                        imgs.Add tmpImg
                    End If
                Next h
            Next j
        End If
    Next i
    
    '读取imgs 中的图像，把图像添加到胶片打印窗口
    For i = 1 To imgs.Count
        imgsPrint.Add imgs(i)
        subChangeLabelForPrint imgsPrint(imgsPrint.Count), 0
    Next i
    
    '图像增加了，调整页数
    Call subRecalPages
    
    '重新显示这一页的图像
    Call subShowPrintImages(Me.VScro.Value)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subDelImage()
'------------------------------------------------
'功能： 删除图像
'       先判断是否有图像被选择了，如果有，则删除被选择的图像，
'       如果没有图像被选择，则删除当前鼠标单击过的图像
'参数：无
'返回：无
'------------------------------------------------
    Dim intStart As Integer
    Dim blnSelected As Boolean      '是否有图像被选择了，如果有则只删除被选择的图像，否则删除当前图像
    Dim i As Integer
    Dim j As Integer
    Dim blnDeleted As Boolean
    Dim intDelViewer As Integer
    Dim intDelImage As Integer
    
    On Error GoTo err
    
    '先判断是否有图像被选择了，如果有，则删除被选择的图像
    blnSelected = False
    For i = Me.FilmViewer.Count - 1 To 1 Step -1
        For j = Me.FilmViewer(i).Images.Count To 1 Step -1
            If funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_选择) = "Select" Then
                intStart = funGetStartImgNo(Me.VScro.Value, i, j)
                imgsPrint.Remove intStart
                blnSelected = True
            End If
        Next
    Next
    
    '如果没有图像被选择，则删除当前鼠标单击过的图像
    If blnSelected = False And Not SelectedImage Is Nothing And mintSelectedViewer <> 0 And mintSelectedImage <> 0 Then
        
        intStart = funGetStartImgNo(Me.VScro.Value, mintSelectedViewer, mintSelectedImage)
        imgsPrint.Remove intStart
        
        intDelViewer = mintSelectedViewer
        intDelImage = mintSelectedImage
        
        blnDeleted = True
    End If
        
    '如果有图像被删除了，则重新显示图像，并且设置下一个被选中的图像，方便连续删除
    If blnDeleted = True Or blnSelected = True Then
        '删除图像之后，图像减少了，重新计算页数
        Call subRecalPages
        '重新显示图像
        Call subShowPrintImages(Me.VScro.Value)
    
        '如果是删除被选中的几个图像，则将选中图像设置成第一张，不需要处理
        '如果是删除当前图像，则需要选中它的前一张图像
        If blnSelected = False Then
            If imgsPrint.Count > 1 Then
                
                intDelImage = intDelImage - 1
                If intDelImage = 0 Then
                    intDelViewer = intDelViewer - 1
                    If intDelViewer <> 0 Then
                        intDelImage = marrPages(Me.VScro.Value).ViewerLayout(intDelViewer).intColumns * marrPages(Me.VScro.Value).ViewerLayout(intDelViewer).intRows
                    Else
                        intDelImage = 0
                    End If
                End If
                
                If intDelViewer > 0 And intDelImage > 0 Then
                    '将选中图像设置成当前图的前一张
                    Call subImageCurrent(1, 1, False)
                    Call subImageCurrent(intDelViewer, intDelImage, True)
                End If
            End If
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subSetFilmFormat()
'------------------------------------------------
'功能：打开排版格式窗口，允许自定义设置胶片的格式
'参数：无
'返回：无
'------------------------------------------------
    Dim i As Integer
    Dim intTmp As Integer
    Dim intPage As Integer                       '''''总页数
    
    On Error GoTo err
    
    Set frmFilmConf.f = Me
    
    '设置排版格式窗口的控件内容
    With frmFilmConf
        .sstabFilmConfig.TabVisible(0) = True
        .sstabFilmConfig.TabVisible(1) = False
        .cobSize.Text = CommBar_Film.FindControl(, ID_frmFilm_FilmSize, , True).Text
        .cobAspect.Text = IIf(mblnIsPortrait, "纵向", "横向")     ''是否纵向打印
        
        If Not mblnIsCustom Then     '标准行列
            .txtRow = UBound(marrRCCount)
            .txtCol = marrRCCount(1)
            .Option(0).Value = True
        Else
            If mblnIsRow Then        '行自定义
                .txtRow = UBound(marrRCCount)
                .txtCol = marrRCCount(1)
                .Option(1).Value = True
            Else                    '列自定义
                .txtCol = UBound(marrRCCount)
                .txtRow = marrRCCount(1)
                .Option(2).Value = True
            End If
            
            For i = 1 To UBound(marrRCCount)
                .txtC(i) = marrRCCount(i)
            Next
        End If
    End With
    
    '显示排版格式窗口
    If frmFilmConf.zlShowMe = True Then
        Call funChangeFormat(Me.VScro.Value, Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text)
'       '重新显示这一页
        Call subShowOnePage(Me.VScro.Value)
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subShiftSelect(intViewerIndex As Integer, intImgIndex As Integer)
'------------------------------------------------
'功能：使用Shift+鼠标左键选择图像
'参数： intViewerIndex --- 当前单击的Viewer索引
'       intImgIndex  --  当前单击的图像索引
'返回：无
'------------------------------------------------
    Dim intStartViewer As Integer
    Dim intStartImage As Integer
    Dim intEndViewer As Integer
    Dim intEndImage As Integer
    Dim i As Integer
    Dim j As Integer
    Dim blnSelected As Boolean
    
    On Error GoTo err
    blnSelected = False
    '先查找第一个被选中的图像
    For i = 1 To FilmViewer.Count - 1
        For j = 1 To FilmViewer(i).Images.Count
            If funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_选择) = "Select" Then
                blnSelected = True
                intStartViewer = i
                intStartImage = j
                Exit For
            End If
        Next j
        If blnSelected = True Then
            Exit For
        End If
    Next i
    
    '如果没有被选中的图像，则只处理当前图像
    If blnSelected = False Then
        Call subImageSelect(intViewerIndex, intImgIndex, True)
    Else
        '如果前面有被选中的图像，则以此图像为第一个图像，循环选择到当前图像
        '判断被选中的图像是在当前图像的前面还是后面
        If intStartViewer < intViewerIndex Then
            intEndViewer = intViewerIndex
            intEndImage = intImgIndex
        ElseIf intStartViewer = intViewerIndex Then
            intEndViewer = intViewerIndex
            If intStartImage <= intImgIndex Then
                intEndImage = intImgIndex
            Else
                intEndImage = intStartImage
                intStartImage = intImgIndex
            End If
        Else
            intEndViewer = intStartViewer
            intEndImage = intStartImage
            intStartViewer = intViewerIndex
            intStartImage = intImgIndex
        End If
        
        '循环选择范围内的图像,其他的图像都不选择
        For i = 1 To FilmViewer.Count - 1
            For j = 1 To FilmViewer(i).Images.Count
                If i = intStartViewer Then
                    If j >= intStartImage And j <= FilmViewer(intStartViewer).Images.Count Then
                        Call subImageSelect(i, j, True)
                    Else
                        Call subImageSelect(i, j, False)
                    End If
                ElseIf i = intEndViewer Then
                    If j >= 1 And j <= intEndImage Then
                        Call subImageSelect(i, j, True)
                    Else
                        Call subImageSelect(i, j, False)
                    End If
                ElseIf i > intStartViewer And i < intEndViewer Then
                    Call subImageSelect(i, j, True)
                Else
                    '不选择
                    Call subImageSelect(i, j, False)
                End If
            Next j
            FilmViewer(i).Refresh
        Next i
    End If
    
    '把选择的图像上传到原始图集合中
    Call subReloadImgsPrint
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub subImageSelect(intViewerIndex As Integer, intImgIndex As Integer, blnSelected As Boolean)
'------------------------------------------------
'功能：选择或者不选择图像
'参数： intViewerIndex --- 需要选择或者取消选择的图像所在的序列
'       intImgIndex --- 需要选择或者取消选择的图像所在的索引
'       blnSelected --- True 选择；False 取消选择
'返回：无
'------------------------------------------------
    On Error GoTo err
    
    If blnSelected = True Then
        Call funSetTagVal(FilmViewer(intViewerIndex).Images(intImgIndex), TAG_选择, "Select")
        FilmViewer(intViewerIndex).Labels(intImgIndex).ForeColour = lngSelectedImageBorderColor
        FilmViewer(intViewerIndex).Labels(intImgIndex).LineStyle = lngSelectedImageBorderLineStyle
        FilmViewer(intViewerIndex).Labels(intImgIndex).LineWidth = lngSelectedImageBorderLineWidth
    Else
        Call funSetTagVal(FilmViewer(intViewerIndex).Images(intImgIndex), TAG_选择, "")
        FilmViewer(intViewerIndex).Labels(intImgIndex).ForeColour = vbWhite
        FilmViewer(intViewerIndex).Labels(intImgIndex).LineStyle = 0
        FilmViewer(intViewerIndex).Labels(intImgIndex).LineWidth = 1
    End If
    FilmViewer(intViewerIndex).Refresh
    Exit Sub
err:
    '暂时不处理

End Sub

Private Sub subImageCurrent(intViewerIndex As Integer, intImgIndex As Integer, blnCurrent As Boolean)
'------------------------------------------------
'功能：设置指定图像是否当前选中的图像
'参数： intViewerIndex --- 需要选择或者取消选择的图像所在的序列
'       intImgIndex --- 需要选择或者取消选择的图像所在的索引
'       blnCurrent --- True 是当前图像；False 不是当前图像
'返回：无
'------------------------------------------------
    On Error GoTo err
    
    If blnCurrent = True Then
        FilmViewer(intViewerIndex).Labels(intImgIndex).ForeColour = lngCurrentImageBorderColor
        FilmViewer(intViewerIndex).Labels(intImgIndex).LineStyle = lngCurrentImageBorderLineStyle
        FilmViewer(intViewerIndex).Labels(intImgIndex).LineWidth = lngCurrentImageBorderLineWidth
        Set SelectedImage = FilmViewer(intViewerIndex).Images(intImgIndex)
        mintSelectedViewer = intViewerIndex
        mintSelectedImage = intImgIndex
    Else
        FilmViewer(intViewerIndex).Labels(intImgIndex).ForeColour = vbWhite
        FilmViewer(intViewerIndex).Labels(intImgIndex).LineStyle = 0
        FilmViewer(intViewerIndex).Labels(intImgIndex).LineWidth = 1
    End If
    FilmViewer(intViewerIndex).Refresh
    Exit Sub
err:
    '暂时不处理

End Sub

Private Sub SelOneSeries()
'------------------------------------------------
'功能：选择当前序列的图像
'参数：
'返回：无
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim strSeriesUID As String
    
    On Error GoTo err
    '先提取当前图像的序列UID
    If SelectedImage Is Nothing Then Exit Sub
    strSeriesUID = SelectedImage.SeriesUID
    
    For i = 1 To FilmViewer.Count - 1
        If FilmViewer(i).Visible = True Then
            For j = 1 To FilmViewer(i).Images.Count
                Call subImageSelect(i, j, IIf(FilmViewer(i).Images(j).SeriesUID = strSeriesUID, True, False))
            Next j
            FilmViewer(i).Refresh
        End If
    Next i
    
    '把选择的图像上传到原始图集合中
    Call subReloadImgsPrint
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SelectInverse()
'------------------------------------------------
'功能：反选图像
'参数：
'返回：无
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo err
    
    If SelectedImage Is Nothing Then Exit Sub
    
    For i = 1 To FilmViewer.Count - 1
        If FilmViewer(i).Visible = True Then
            For j = 1 To FilmViewer(i).Images.Count
                If funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_选择) = "Select" Then
                    Call subImageSelect(i, j, False)
                Else
                    Call subImageSelect(i, j, True)
                End If
            Next j
            FilmViewer(i).Refresh
        End If
    Next i
    
    '把选择的图像上传到原始图集合中
    Call subReloadImgsPrint
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub subOpenFilmView()
'------------------------------------------------
'功能：打开图像处理窗口
'参数： thisViewer - 新打开窗口的Viewer,主要是提取宽度和高度
'返回：无
'------------------------------------------------
    Dim dcmNewImage As DicomImage
    
    On Error GoTo err
    
    '如果已经打开了一个图像处理窗体，则不再打开另一个
    If Not mfrmFilmView Is Nothing Then Exit Sub
    If SelectedImage Is Nothing Then Exit Sub
    If mintSelectedViewer = 0 Then Exit Sub
    If mintSelectedImage = 0 Then Exit Sub
    If FilmViewer(mintSelectedViewer).Images.Count < mintSelectedImage Then Exit Sub
    
    Set mfrmFilmView = New frmFilmView
    
    Call mfrmFilmView.zlShowMe(SelectedImage, Me, mintSelectedViewer, mintSelectedImage)
    
    '    挂上截获消息的hook，不能放在mfrmFilm的load事件
    plngFilmViewPreWndProc = FilmViewHook(mfrmFilmView.hwnd)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub subFilmViewButtonClick(control As CommandBarControl)
'------------------------------------------------
'功能：打开图像处理窗口
'参数： thisViewer - 新打开窗口的Viewer,主要是提取宽度和高度
'返回：无
'------------------------------------------------
    On Error GoTo err
    
    If mfrmFilmView Is Nothing Then Exit Sub
    
    Call mfrmFilmView.ZLToolButtonClick(control)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub subCutOutRatio(lngControlID As Long)
'------------------------------------------------
'功能：设置裁剪的固定比例，并画出固定比例的裁剪框
'参数： lngControlID --- 裁剪菜单项ID
'返回：无
'------------------------------------------------
    On Error GoTo err
    Dim intLeft As Integer
    Dim intTop As Integer
    Dim intWidth As Integer
    Dim intHeight As Integer
    Dim dblScaleRatio As Double
    
    If SelectedImage Is Nothing Then Exit Sub
    If mintSelectedViewer = 0 Then Exit Sub
    If mintSelectedImage = 0 Then Exit Sub
    
    '设置裁剪的比例
    Select Case lngControlID
        Case ID_frmFilm_CutOut_14X17
            mdblCutOutRatio = 14 / 17
            dblScaleRatio = 1
        Case ID_frmFilm_CutOut_11X14
            mdblCutOutRatio = 11 / 14
            dblScaleRatio = 11 / 14
        Case ID_frmFilm_CutOut_10X14
            mdblCutOutRatio = 10 / 14
            dblScaleRatio = 10 / 14
        Case ID_frmFilm_CutOut_8X10
            mdblCutOutRatio = 8 / 10
            dblScaleRatio = 8 / 14
        Case ID_frmFilm_CutOut_14X14
            mdblCutOutRatio = 14 / 14
            dblScaleRatio = 1
        Case ID_frmFilm_CutOut_17X14
            mdblCutOutRatio = 17 / 14
            dblScaleRatio = 1
        Case ID_frmFilm_CutOut_14X11
            mdblCutOutRatio = 14 / 11
            dblScaleRatio = 14 / 17
        Case ID_frmFilm_CutOut_14X10
            mdblCutOutRatio = 14 / 10
            dblScaleRatio = 14 / 17
        Case ID_frmFilm_CutOut_10X8
            mdblCutOutRatio = 10 / 8
            dblScaleRatio = 10 / 17
    End Select
    
    '显示固定比例的裁剪框
    If mintSelectedViewer < FilmViewer.Count Then
        If mintSelectedImage <= FilmViewer(mintSelectedViewer).Images.Count Then
            '根据SelectImg计算裁剪框的位置
            intHeight = SelectedImage.sizey
            intWidth = intHeight * mdblCutOutRatio
            If intWidth > SelectedImage.sizex Then
                '交换宽度和高度
                intWidth = SelectedImage.sizex
                intHeight = intWidth / mdblCutOutRatio
            End If
            
            '根据缩放比例重新计算宽度和高度
            intWidth = intWidth * dblScaleRatio
            intHeight = intHeight * dblScaleRatio
            
            '计算Left和Top
            intTop = (SelectedImage.sizey - intHeight) / 2
            intLeft = (SelectedImage.sizex - intWidth) / 2
            
            '显示裁剪框
            FilmViewer(mintSelectedViewer).Images(mintSelectedImage).Labels.Add GetNewLabel(doLabelRectangle, intLeft, intTop, intWidth, intHeight)
            Set mdcmSelectLabel = FilmViewer(mintSelectedViewer).Images(mintSelectedImage).Labels(FilmViewer(mintSelectedViewer).Images(mintSelectedImage).Labels.Count)
            mdcmSelectLabel.Tag = CUT_LABEL
            mintCutOutViewer = mintSelectedViewer
            mintCutOutImage = mintSelectedImage
            mintCutOutLabel = FilmViewer(mintSelectedViewer).Images(mintSelectedImage).Labels.Count
            FilmViewer(mintSelectedViewer).Refresh
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub subCutOutClick()
'------------------------------------------------
'功能：单击裁剪按钮，自动触发下拉列表中对应被选中的项目
'参数：
'返回：无
'------------------------------------------------
    On Error GoTo err
    '判断当前是那种裁剪方式
    If CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_Custom, , True).Checked Then
        mdblCutOutRatio = 0
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_14X17, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_14X17)
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_11X14, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_11X14)
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_10X14, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_10X14)
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_8X10, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_8X10)
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_14X14, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_14X14)
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_17X14, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_17X14)
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_14X11, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_14X11)
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_14X10, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_14X10)
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_10X8, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_10X8)
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub subCutOutButtonState(lngControlID As Long)
'------------------------------------------------
'功能：调整裁剪子菜单的状态
'参数：
'返回：无
'------------------------------------------------
    On Error GoTo err
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_Custom, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_14X17, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_11X14, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_10X14, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_8X10, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_14X14, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_17X14, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_14X11, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_14X10, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_10X8, , True).Checked = False
        
        CommBar_Film.Item(3).FindControl(, lngControlID, , True).Checked = True
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut, , True).IconId = lngControlID
        
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub subCalImageMaxSize(strFilmSize As String, strFormat As String, intImageResolution As Integer, _
    arrImageSize() As ImageSize)
'------------------------------------------------
'功能：根据胶片尺寸，图像布局，计算每一个图像框的最大尺寸
'参数： strFilmSize --- 胶片尺寸
'       strFormat --- 胶片格式
'       intImageResolution --- 相机的图像分辨率
'       arrImageSize --- [OUT]返回每个图像框的最大分辨率
'返回：无
'------------------------------------------------
    Dim intImageCount As Integer
    Dim lngFilmWidth As Long
    Dim lngFilmHeight As Long
    Dim strCurFormat As String
    Dim i As Integer
    Dim j As Integer
    Dim intX As Integer
    Dim intY As Integer
    Dim intCount As Integer
    
    ReDim arrImageSize(0)
    
    On Error GoTo err
    
    '解析胶片的尺寸，有两种情况：“8_5INX11IN”，“14INX17IN”
    If UBound(Split(UCase(strFilmSize), "X")) = 1 Then
        lngFilmWidth = Val(Replace(Split(UCase(strFilmSize), "X")(0), "_", ".")) * intImageResolution
        lngFilmHeight = Val(Replace(Split(UCase(strFilmSize), "X")(1), "_", ".")) * intImageResolution
    Else
        '胶片尺寸不正确，退出解析
        Exit Sub
    End If
    
    '根据格式解析图像的数量，有三种格式表示方法：“STANDARD\1,2”（行，列），“ROW\1,2”，“COL\2,3”，分别解析
    If InStr(UCase(strFormat), "STANDARD\") > 0 Then
        strCurFormat = Mid(strFormat, 10)
        If UBound(Split(strCurFormat, ",")) = 1 Then
            intX = Val(Split(strCurFormat, ",")(0))
            intY = Val(Split(strCurFormat, ",")(1))
            ReDim arrImageSize(intX * intY)
            For i = 1 To intY
                For j = 1 To intX
                    arrImageSize((i - 1) * intX + j).intWidth = lngFilmWidth / intX
                    arrImageSize((i - 1) * intX + j).intHeight = lngFilmHeight / intY
                Next j
            Next i
        Else
            Exit Sub
        End If
    ElseIf InStr(UCase(strFormat), "ROW\") > 0 Then
        strCurFormat = Mid(strFormat, 5)
        intY = UBound(Split(strCurFormat, ",")) + 1
        intCount = 0
        For i = 1 To intY
            intX = Val(Split(strCurFormat, ",")(i - 1))
            ReDim Preserve arrImageSize(UBound(arrImageSize) + intX)
            For j = 1 To intX
                intCount = intCount + 1
                arrImageSize(intCount).intWidth = lngFilmWidth / intX
                arrImageSize(intCount).intHeight = lngFilmHeight / intY
            Next j
        Next i
    ElseIf InStr(UCase(strFormat), "COL\") > 0 Then
        strCurFormat = Mid(strFormat, 5)
        intX = UBound(Split(strCurFormat, ",")) + 1
        intCount = 0
        For i = 1 To intX
            intY = Val(Split(strCurFormat, ",")(i - 1))
            ReDim Preserve arrImageSize(UBound(arrImageSize) + intY)
            For j = 1 To intY
                intCount = intCount + 1
                arrImageSize(intCount).intWidth = lngFilmWidth / intX
                arrImageSize(intCount).intHeight = lngFilmHeight / intY
            Next j
        Next i
    Else
        '胶片布局不正确，退出解析
        Exit Sub
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subFillPageRCCount(strFilmFormat As String)
'------------------------------------------------
'功能：解析界面布局串，填写界面布局数组和对应参数
'参数： strFilmFormat界面布局串
'返回：无
'------------------------------------------------
    Dim i As Integer
    Dim strFormatType As String
    Dim strFormatNumber As String
    Dim strRCDetail() As String
    Dim intRCCount As Integer
    
    On Error GoTo err
    
    '解析胶片布局
    If strFilmFormat <> "" Then
        i = InStr(strFilmFormat, "\")
        strFormatType = Mid(strFilmFormat, 1, i - 1)
        strFormatNumber = Mid(strFilmFormat, i + 1, Len(strFilmFormat) - i)
    End If
    
    '判断是行优先？列优先？标准格式？
    If strFormatType = "ROW" Then
        mblnIsRow = True
        mblnIsCustom = True
    ElseIf strFormatType = "COL" Then
        mblnIsRow = False
        mblnIsCustom = True
    ElseIf strFormatType = "STANDARD" Then
        mblnIsRow = True
        mblnIsCustom = False
    End If

    '提取布局中的具体行列数值，保存到aRCCount数组中
    strRCDetail = Split(strFormatNumber, ",")
    If UBound(strRCDetail) >= 0 Then
        If Not mblnIsCustom Then                     '标准定义行列数
            intRCCount = Val(strRCDetail(1))
            ReDim marrRCCount(intRCCount)              '''每行/每列的图像数目
            For i = 1 To UBound(marrRCCount)
                marrRCCount(i) = Val(strRCDetail(0))
            Next
        Else
            intRCCount = UBound(strRCDetail) + 1            ''''行数或列数目
            ReDim marrRCCount(intRCCount)              '''每行/每列的图像数目
            For i = 1 To UBound(marrRCCount)
                marrRCCount(i) = Val(strRCDetail(i - 1))
            Next
        End If
    End If
        
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subImageSort(blnIncrease As Boolean)
'------------------------------------------------
'功能：对当前胶片中，选中的图像进行排序，按照图像号排序
'参数： blnIncrease -正序
'返回：无
'------------------------------------------------
    Dim v As DicomViewer
    Dim i As Integer
    Dim j As Integer
    Dim intSelectedCount As Integer
    Dim blnSortSelected As Boolean  '只对被选中的图像进行排序
    Dim intCurrImgIndex As Integer
    Dim intImgsPrintStartIndex As Integer   '开始排序的正本图像Index
    Dim intImgsPrintEndIndex As Integer     '结束排序的正本图像Index
    
    On Error GoTo err
    
    '首先将当前的Select状态更新到正本中
    Call subReloadImgsPrint
    
    '如果当前胶片中有被选中的2个以上图片，则对被选中的图像进行排序，否则对胶片中的所有图像进行排序
    intSelectedCount = 0
    For Each v In FilmViewer
        For i = 1 To v.Images.Count
            If funGetTagVal(v.Images(i).Tag, TAG_选择) = "Select" Then
                intSelectedCount = intSelectedCount + 1
                If intSelectedCount >= 2 Then
                    blnSortSelected = True
                    Exit For
                End If
            End If
        Next i
        If blnSortSelected = True Then Exit For
    Next v
    
    '对图像进行排序
    intImgsPrintStartIndex = funGetStartImgNo(VScro.Value, 1, 1)
    intImgsPrintEndIndex = funGetStartImgNo(VScro.Value, FilmViewer.Count - 1, 1) + marrPages(VScro.Value).ViewerLayout(FilmViewer.Count - 1).intColumns * marrPages(VScro.Value).ViewerLayout(FilmViewer.Count - 1).intRows - 1
    If intImgsPrintEndIndex > imgsPrint.Count Then intImgsPrintEndIndex = imgsPrint.Count
    
    '在正本图像中进行排序
    For i = intImgsPrintStartIndex To intImgsPrintEndIndex
        If (blnSortSelected = True And funGetTagVal(imgsPrint(i).Tag, TAG_选择) = "Select") Or blnSortSelected = False Then
            '开始往后比对，调整这个图像的位置
            Call subImageSortAndMove(i, intImgsPrintEndIndex, blnSortSelected, blnIncrease)
        End If
    Next i
    
    '全部排序完成之后，重新显示
    '调用过程调整图像的显示
    Call subShowPrintImages(Me.VScro.Value)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subImageSortAndMove(intImgsPrintIndex As Integer, intImgsPrintEndIndex As Integer, _
    blnSortSelected As Boolean, blnIncrease As Boolean)
'------------------------------------------------
'功能：从intImageIndex开始，向后查找排序，调整这个图像的位置
'参数： intImgsPrintIndex -当前开始查找的图像index，原图中的index
'       intImgsPrintEndIndex -- 在正本图像中查找的结束Index
'       blnSortSelected -- 只对被选中的图像进行排序
'       blnIncrease -- True 升序排序，False 逆序排序
'返回：无
'------------------------------------------------
    Dim intNextImageIndex As Integer
    Dim lngCurrImgNum As Long       '记录当前图像的图像号
    Dim lngTestingImgNum As Long    '记录被测试的图像的图像号
    Dim v As DicomViewer
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo err
    
    '记录当前图像的图像号
    lngTestingImgNum = 0
    intNextImageIndex = 0
    If imgsPrint(intImgsPrintIndex).Attributes(&H20, &H13).Exists And Not IsNull(imgsPrint(intImgsPrintIndex).Attributes(&H20, &H13).Value) Then
        lngCurrImgNum = Val(imgsPrint(intImgsPrintIndex).Attributes(&H20, &H13).Value)
    End If
    
    For i = intImgsPrintIndex + 1 To intImgsPrintEndIndex
        If (blnSortSelected = True And funGetTagVal(imgsPrint(i).Tag, TAG_选择) = "Select") Or blnSortSelected = False Then
            '提取图像号
            If imgsPrint(i).Attributes(&H20, &H13).Exists And Not IsNull(imgsPrint(i).Attributes(&H20, &H13).Value) Then
                lngTestingImgNum = Val(imgsPrint(i).Attributes(&H20, &H13).Value)
            End If

            If (blnIncrease = True And lngTestingImgNum < lngCurrImgNum And lngTestingImgNum <> 0) _
                Or (blnIncrease = False And lngTestingImgNum > lngCurrImgNum And lngTestingImgNum <> 0) Then
                intNextImageIndex = i
                lngCurrImgNum = lngTestingImgNum
            End If
        End If
    Next i
    
    '移动图像
    If intNextImageIndex <> 0 Then
        Call imgsPrint.Move(intNextImageIndex, intImgsPrintIndex)
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subShowPrintImages(intPage As Integer)
'------------------------------------------------
'功能：根据胶片格式的行数和列数，重新显示指定页的图像，并显示图像相关的定位线等
'参数：
'返回：无
'------------------------------------------------

    On Error GoTo err
    
     '前提条件，Viewer的数量和位置，在其他过程中已经处理好了
     
    If Not mblnBegin Then Exit Sub
    
    '加载显示一页的图像
    Call subLoadPrintImage(intPage)
    
    '显示定位线
    Call subDispReferLineFilm
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub


Public Function ZLAddImage(img As DicomImage, blnPrinted As Boolean, dblWidth As Double, dblHeight As Double) As Long
'------------------------------------------------
'功能：增加一个图像
'参数： img -- 需要增加的图像
'       blnPrinted -- 记录图像是否已经被打印过
'       dblWidth --  图像原来显示时占用的宽度，用来调整图像的缩放比例和移动
'       dblHeight -- 图像原来显示时占用的高度，用来调整图像的缩放比例和移动
'返回：0 -- 正确；1--出错
'------------------------------------------------
    Dim AddedImage As DicomImage
    Dim dblScale As Double
    Dim thisViewer As DicomViewer
    
    On Error GoTo err
    
    imgsPrint.Add img
    Set AddedImage = imgsPrint(imgsPrint.Count)
    
    '处理打印标记
    If blnPrinted = True Then mblnPrinted = True
    
    '处理图像的标注
    If AddedImage.Labels(G_INT_SYS_LABEL_PAT_INFO).Visible = False Then
        Call subDispImageInfo(AddedImage, True, False, True)        ''显示病人四角信息和窗宽窗位信息
    End If
    
    '处理放大后的图像的显示
    Set thisViewer = FilmViewer(FilmViewer.Count - 1)
    
    '应该统一调用 subScaleImage 过程，但是这个过程图像显示会歪，先不使用
'    Call subScaleImage(AddedImage, thisViewer, CLng(dblWidth), CLng(dblHeight))
    
    dblScale = ((thisViewer.width / thisViewer.MultiColumns) / dblWidth + (thisViewer.height / thisViewer.MultiRows) / dblHeight) / 2
    AddedImage.Zoom = AddedImage.Zoom * dblScale
    AddedImage.ScrollX = AddedImage.ScrollX * dblScale
    AddedImage.ScrollY = AddedImage.ScrollY * dblScale
    
    '增加了图像，调整参数
    Call subChangeLabelForPrint(AddedImage, 0)
    '图像增加了，调整页数
    Call subRecalPages
    
    '如果新添加的图像，在当前页，则重新显示这一页的图像
    If imgsPrint.Count < funGetStartImgNo(Me.VScro.Value, 1, 1) + marrPages(Me.VScro.Value).intImageCount Then
        Call subShowPrintImages(Me.VScro.Value)
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    ZLAddImage = 1
End Function


Private Function funChangeFormat(intPage As Integer, strPageFormat As String, Optional intViewerIndex As Integer = 0, _
    Optional intRows As Integer = 0, Optional intCols As Integer = 0) As Long
'------------------------------------------------
'功能：更改页面的显示布局，包括胶片布局和组合图像布局
'参数： intPage --      需要调整的页数
'       strPageFormat --页面中的Viewer布局，DICOM格式串
'       intViewerIndex--【可选】图像组合时用，0-不做图像组合；其他-进行图像组合的Viewer号
'       intRows --      【可选】图像组合时用，0-不做图像组合；其他-进行图像组合的行数
'       intCols --      【可选】图像组合时用，0-不做图像组合；其他-进行图像组合的列数
'返回：0 -- 正确；1--出错
'------------------------------------------------
    Dim intCurrViewerCount As Integer
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo err
    
    funChangeFormat = 1
    If intPage > UBound(marrPages) Then Exit Function
    If intViewerIndex <> 0 Then
        If intRows = 0 Then intRows = 1
        If intCols = 0 Then intCols = 1
    End If
    
    '先记录Viewer数量和页面布局
    intCurrViewerCount = funGetPageViewerCount(strPageFormat)
    
    '调整当前页面中的图像组合布局
    If intViewerIndex <> 0 And intViewerIndex <= intCurrViewerCount Then
        marrPages(intPage).ViewerLayout(intViewerIndex).intColumns = intCols
        marrPages(intPage).ViewerLayout(intViewerIndex).intRows = intRows
    End If
    
    '调整当前页面和后续页面的页面布局
    For i = intPage To UBound(marrPages)
        '填写这一页的Viewer数量和图像数量
        marrPages(i).intViewerCount = intCurrViewerCount
        marrPages(i).strPageFormat = strPageFormat
        
        'Viewer数量增加或者减少了，需要调整这个页面中的ViewerLayout数组,确保新增的行数，列数=1
        ReDim Preserve marrPages(i).ViewerLayout(intCurrViewerCount)
        
        For j = 1 To intCurrViewerCount
            '设置图像组合的布局
            If marrPages(i).ViewerLayout(j).intColumns = 0 Then marrPages(i).ViewerLayout(j).intColumns = 1
            If marrPages(i).ViewerLayout(j).intRows = 0 Then marrPages(i).ViewerLayout(j).intRows = 1
        Next j
        '计算这一页的图像总数，不是实际图像总数，是这一页总共能够摆放的图像总数
        marrPages(i).intImageCount = funGetPageImageCount(i)
    Next i
    
    '根据图像正本数量，布局情况，重新计算胶片页数
    Call subRecalPages
    
    funChangeFormat = 0
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub subRecalPages()
'------------------------------------------------
'功能：根据正本图像的数量，重新计算实际页数
'参数：
'返回：
'------------------------------------------------
    Dim intPageCount As Integer
    Dim intImageCount As Integer
    Dim intDefaultViewerCount As Integer
    Dim strDefaultFormat As String
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo err
    
    '根据图像总数，调整全部页面布局
    intPageCount = 0
    intImageCount = 0
    
    '最后一页的参数，作为默认参数
    intDefaultViewerCount = marrPages(UBound(marrPages)).intViewerCount
    strDefaultFormat = marrPages(UBound(marrPages)).strPageFormat
    
    '如果当前一张图都没有，页面设置成一页
    ReDim Preserve marrPages(IIf(imgsPrint.Count > 0, imgsPrint.Count, 1))
    
    For i = 1 To imgsPrint.Count
        '处理新增的页面，将当前默认格式应用到后续页面，只应用页面布局，不应用图像组合布局
        If marrPages(i).intViewerCount = 0 Then
            '填写这一页的Viewer数量和图像数量
            marrPages(i).intViewerCount = intDefaultViewerCount
            marrPages(i).strPageFormat = strDefaultFormat
            
            'Viewer数量增加或者减少了，需要调整这个页面中的ViewerLayout数组,确保新增的行数，列数=1
            ReDim Preserve marrPages(i).ViewerLayout(intDefaultViewerCount)
            
            For j = 1 To intDefaultViewerCount
                marrPages(i).ViewerLayout(j).intColumns = 1
                marrPages(i).ViewerLayout(j).intRows = 1
            Next j
            
            '计算这一页的图像总数，不是实际图像总数，是这一页总共能够摆放的图像总数
            marrPages(i).intImageCount = funGetPageImageCount(i)
        End If
        
        intPageCount = intPageCount + 1
        intImageCount = intImageCount + marrPages(i).intImageCount
        
        '如果累加的图像总数超过正本图像总数，则退出循环，得到当前的胶片页面数
        If intImageCount >= imgsPrint.Count Then Exit For
    Next i
    
    ReDim Preserve marrPages(IIf(intPageCount > 0, intPageCount, 1))
    
    '设置滚动条的最大值
    Me.VScro.Max = UBound(marrPages)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function funGetPageViewerCount(strPageFormat As String) As Integer
'------------------------------------------------
'功能：根据页面布局方式，计算当前页面的Viewer总数
'参数： strPageFormat -- DICOM标准的页面布局
'返回：无
'------------------------------------------------
    Dim i As Integer
    Dim strFormatType As String
    Dim strFilmFormat As String
    Dim strRCDetail() As String
    
    funGetPageViewerCount = 0
    
    On Error GoTo err
    
    If strPageFormat = "" Then Exit Function
    
    '解析胶片的页面布局
    i = InStr(strPageFormat, "\")
    strFormatType = Mid(strPageFormat, 1, i - 1)
    strFilmFormat = Mid(strPageFormat, i + 1, Len(strPageFormat) - i)
    
    strRCDetail = Split(strFilmFormat, ",")
    If UBound(strRCDetail) >= 0 Then
        If strFormatType = "STANDARD" Then
            '标准布局，直接=行数*列数
            funGetPageViewerCount = Val(strRCDetail(0)) * Val(strRCDetail(1))
        Else
            '异性布局，行优先，或者列优先，逐个相加
            For i = 0 To UBound(strRCDetail)
                funGetPageViewerCount = funGetPageViewerCount + Val(strRCDetail(i))
            Next i
        End If
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    funGetPageViewerCount = 0
End Function

Private Function funGetPageImageCount(intPage As Integer) As Integer
'------------------------------------------------
'功能：根据页面布局方式，计算当前页面能容纳的图像总数
'参数： intPage -- 需要计算的页数
'返回：当前页面能容纳的图像总数
'------------------------------------------------
    Dim i As Integer
    Dim intImageCount As Integer
    
    On Error GoTo err
    
    funGetPageImageCount = 0
    If intPage > UBound(marrPages) Then Exit Function
    
    intImageCount = 0
    For i = 1 To marrPages(intPage).intViewerCount
        intImageCount = intImageCount + (marrPages(intPage).ViewerLayout(i).intColumns * marrPages(intPage).ViewerLayout(i).intRows)
    Next i
    
    funGetPageImageCount = intImageCount
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    
End Function

Private Sub InitPageFormat(strPageFormat As String)
'------------------------------------------------
'功能：初始化页面的布局设置
'参数：strPageFormat -- DICOM格式的页面布局
'返回：无
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
    
    '一张图都没有的时候，marrPages设置成1页
    
    ReDim marrPages(1)
    marrPages(1).strPageFormat = strPageFormat

    '重新计算这一页中的Viewer数量
    marrPages(1).intViewerCount = funGetPageViewerCount(strPageFormat)
    ReDim marrPages(1).ViewerLayout(marrPages(1).intViewerCount)
    
    '每一个Viewer都设置成一行X一列
    For i = 1 To marrPages(1).intViewerCount
        marrPages(1).ViewerLayout(i).intColumns = 1
        marrPages(1).ViewerLayout(i).intRows = 1
    Next i
    
    '填写这一页的最大图像数量
    marrPages(1).intImageCount = funGetPageImageCount(1)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function funGetTagVal(strTag As String, TagID As String) As String
'------------------------------------------------
'功能：从TAG中提取对应的值，目前TAG总共包含4个值
'参数： strTag --- 需要提取的TAG
'       TagID --- TAG的索引，使用定义好的常量
'返回：TagID对应的值
'------------------------------------------------
    Dim arrTags() As String
    Dim intTagID As Integer
    
    On Error GoTo err
    intTagID = Val(TagID)
    If intTagID <= 0 Or intTagID > 4 Then Exit Function
    If strTag = "" Then Exit Function
    
    arrTags = Split(strTag, zlSpliter)
    If UBound(arrTags) = 4 Then
        funGetTagVal = arrTags(intTagID)
    End If
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function funSetTagVal(dcmImage As DicomImage, TagID As String, strTagVal As String) As Boolean
'------------------------------------------------
'功能：从TAG中提取对应的值，目前Tag总共包含4个值
'参数： dcmImage --- 需要提取的TAG的DICOM图像
'       TagID --- TAG的索引，使用定义好的常量
'       strTagVal --- 要设置的值
'返回：True成功，False 失败
'------------------------------------------------
    
    Dim strTag As String
    Dim arrTags() As String
    Dim intTagID As Integer
    Dim i As Integer
    
    On Error GoTo err
    
    If dcmImage Is Nothing Then Exit Function
    
    strTag = dcmImage.Tag
    intTagID = Val(TagID)
    If intTagID <= 0 Or intTagID > 4 Then Exit Function
    arrTags = Split(strTag, zlSpliter)
    
    If UBound(arrTags) < intTagID Then
        ReDim Preserve arrTags(intTagID) As String
    End If
    arrTags(intTagID) = strTagVal
    
    strTag = ""
    For i = 1 To UBound(arrTags)
        strTag = strTag & zlSpliter & arrTags(i)
    Next i
    If i <= 4 Then
        For i = i To 4
            strTag = strTag & zlSpliter & ""
        Next i
    End If
    
    dcmImage.Tag = strTag
    funSetTagVal = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub subShowOnePage(intPage As Integer)
'------------------------------------------------
'功能：显示当前页面的图像，适用于图像总数没变，但是页面或者布局改变的情况
'参数： intPage -- 需要显示的页数
'返回： 无
'------------------------------------------------

    On Error GoTo err
    
    '根据页面布局，加载页面中的Viewer
    Call subLoadViewer(intPage)

    '调用过程调整图像的显示
    Call subShowPrintImages(intPage)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub MouseWheel(intDirection As Integer)
'------------------------------------------------
'功能：处理鼠标滚轮的消息
'参数：intDirection--滚轮滚动方向 1-鼠标上滚；2-鼠标下滚
'返回：无
'------------------------------------------------
    Dim dblScale As Double
    
    '发生错误，不做任何提示
    On Error Resume Next
    
    If SelectedImage Is Nothing Then Exit Sub
    If mintSelectedViewer > FilmViewer.Count Then Exit Sub
    If FilmViewer.Count = 1 Then Exit Sub
    
    If intDirection = 1 Then
        dblScale = 1 + (lngZoomStep * 0.01)
    Else
        dblScale = 1 - (lngZoomStep * 0.01)
    End If
    
    Call subCenterZoom(SelectedImage, FilmViewer(mintSelectedViewer), SelectedImage.ActualZoom * dblScale)
        
    If SelectedImage.Labels.Count >= G_INT_SYS_LABEL_RULLER Then
        If SelectedImage.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then  '更新标尺单位
            Call UpdateRuler(SelectedImage, True)
        End If
    End If
    
    '序列同步
    Call subSynchronalImg(False, IMG_SYN_ZOOMPAN)
End Sub

