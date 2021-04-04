VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "*\Azl9PacsControl\zl9PacsControl.vbp"
Begin VB.Form frmImageProcessV2 
   Caption         =   "图像处理"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11550
   Icon            =   "frmImageProcessV2.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin zl9PacsControl.ucSplitter ucSplitter 
      Height          =   6375
      Left            =   4215
      TabIndex        =   8
      Top             =   480
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   11245
      Control1Name    =   "ucBgImages"
      Control2Name    =   "DViewer"
   End
   Begin zl9PACSWork.ucBgImgViewer ucBgImages 
      Height          =   6375
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   11245
   End
   Begin VB.ListBox lstMemoText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   8280
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox picMemo 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11055
      TabIndex        =   2
      Top             =   7080
      Width           =   11055
      Begin VB.PictureBox picCboDropDown 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   8040
         Picture         =   "frmImageProcessV2.frx":6852
         ScaleHeight     =   375
         ScaleWidth      =   255
         TabIndex        =   10
         Top             =   120
         Width           =   255
      End
      Begin VB.ComboBox cbxMemoText 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         TabIndex        =   9
         Top             =   120
         Width           =   5655
      End
      Begin VB.CommandButton cmdFont 
         Height          =   375
         Left            =   9000
         Picture         =   "frmImageProcessV2.frx":6BAE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "设置当前备注字体。"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdInsert 
         Height          =   375
         Left            =   8640
         Picture         =   "frmImageProcessV2.frx":6EF0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "将当前备注设置为常用备注"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   375
         Left            =   8280
         Picture         =   "frmImageProcessV2.frx":765A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "添加备注"
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblMemoText 
         AutoSize        =   -1  'True
         Caption         =   "添加备注文字："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   6
         Top             =   195
         Width           =   1470
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   1800
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   0
   End
   Begin MSComDlg.CommonDialog diaFont 
      Left            =   2520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtInputText 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   7800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin DicomObjects.DicomViewer DViewer 
      Height          =   6375
      Left            =   4350
      TabIndex        =   7
      Top             =   480
      Width           =   6255
      _Version        =   262147
      _ExtentX        =   11033
      _ExtentY        =   11245
      _StockProps     =   35
      BackColor       =   0
      UseScrollBars   =   0   'False
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmImageProcessV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TPoint
  X As Integer
  Y As Integer
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type


'Private gobjImageProcess As frmImageProcess

Private glngColor(10) As Long             '标记图中圆形编号使用的9个颜色

Private Const G_STR_TAG = "Po=息肉[+]E=糜烂区[+]M=镶嵌[+]L=粘膜白斑[+]C=湿疣[+]I=浸润性癌[+]W=醋酸白色上皮[+]AT=异常转化区[+]V=非典型血管[+]P=点状血管[+]Xn=直接活检部位"

'图像处理
Private Const conMenu_Process_Window = 501           '亮度对比度
Private Const conMenu_Process_Zoom = 502             '缩放
Private Const conMenu_Process_Corp = 512             '拖动
Private Const conMenu_Process_RRotate = 503          '顺时针旋转
Private Const conMenu_Process_LRotate = 504          '逆时针旋转
Private Const conMenu_Process_Sharpness = 505        '锐化
Private Const conMenu_Process_Filter = 506           '平滑
Private Const conMenu_Process_Arrow = 507            '箭头标注
Private Const conMenu_Process_Ellipse = 508          '圆形标注
Private Const conMenu_Process_Text = 509             '文字标注
Private Const conMenu_Process_RectZoom = 510         '裁剪采集
Private Const conMenu_Process_RectCapture = 511      '裁剪后采集
Private Const conMenu_Process_Line = 520             '直线标注
Private Const conMenu_Process_Exit = 2613            '退出
Private Const conMenu_Process_Save = 3091            '保存
Private Const conMenu_Process_SaveToReport = 3941    '保存到检查
Private Const conMenu_Process_SaveToStudy = 3943     '保存到报告
Private Const conMenu_Process_DelAllLabels = 8113    '删除全部标注，使用其他系统的图标编号
Private Const conMenu_Process_MoveLabel = 6891       '移动或删除选中标注，使用其他系统的图标编号
Private Const conMenu_Process_LabelSetUp = 10003     '标注按钮设置，使用其他系统的图标编号
Private Const conMenu_Process_Restore = 8124         '恢复
Private Const conMenu_Process_TextTag = 5010         '文本标记
Private Const conMenu_Process_NumTag = 7405          '数字标记
Private Const conMenu_Process_Page = 1001
Private Const conMenu_Process_Num = 96
Private Const conMenu_Process_Word = 97



'Private mlngModule As Long
Private mlngAdviceId As Long

Private mImage As DicomImage
Private mintMouseState As TMouseState
Private mblnDcmViewDown As Boolean
Private mMouseDownPoint As TPoint
Private mInitScrollPoint As TPoint
Private mCorpSize As TPoint             '拖动后的相对偏移位置

'调窗和漫游使用的鼠标基准位置
Private mlngBaseXX As Long
Private mlngBaseYY As Long
'移动标注使用的鼠标基准位置
Private mlngBaseX As Long
Private mlngBaseY As Long

Private mdcmSelectLabel As DicomLabel   '当前被选中的标注
Private mMovingLabel As DicomLabel      '当前选中要移动或者删除的标注

Private mblnOk As Boolean
Private mOldImage As DicomImage
Private mlngImgIndex As Long            '父窗体选中缩略图的索引
Private mintTextIndex As Integer        '文字标注按钮的索引
Private mstrText As String              '文字标注内容
Private mstrCustom  As String           '自定义标注内容
Private mintNumberIndex As Integer      '数字编号按钮的索引
Private mintAutoNumber As Integer       '自动递增编号的最大号码
Private mStrTemp As String
Private mstrUser As String
 
Private mlngWinType As TImgProcessType  '打开窗口时窗口类型

Private mlngPreViewTime As Long         '移动预览延时关闭时间
Private mlngState As Long               '预览图像窗口状态，1-预览；2-处理；3-单击后
Private mblnMoved As Boolean
Private mstrQueryValue As String
Private mblnIsUnloud As Boolean         '当前鼠标位置是否自动关闭
Private mblnDrag As Boolean
Private mintDisState As Integer
Private mblnIsChanged As Boolean
Private mblnCase As Boolean
Private mblnIsReportShow As Boolean
Private maryImgInfos() As Object

Private mrsTmp As ADODB.Recordset       '图像备注记录集

Private Enum TMouseState
    msNone = 0          '无状态
    msWinLevel = 1      '窗宽窗位
    msZoom = 2          '缩放
    msRectangle = 3     '框选缩放
    msline = 10         '直线
    msArrow = 11        '箭头
    msEllipse = 12      '椭圆
    msText = 13         '文字
    msDrag = 14         '漫游拖动
    msNumber = 15       '数字编号
    msFixText = 16      '文字按钮
    msMove = 17         '移动和删除标注
End Enum


Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Event OnUnload()
Public Event OnSaveImage(ByVal emImageType As TImageType, ByRef dcmImage As DicomImage)

Private mblnAllowSaveStudyImg As Boolean    '允许保存检查图
Private mblnAllowSaveReportImg As Boolean   '允许保存报告图


Public Sub SetButtonState(ByVal blnStudyImgSaveState As Boolean, ByVal blnRepImgSaveState As Boolean)
    mblnAllowSaveStudyImg = blnStudyImgSaveState
    mblnAllowSaveReportImg = blnRepImgSaveState
End Sub



Property Get WinType() As Long
    WinType = mlngWinType
End Property

Private Sub LoadImgs(objImgInfos() As Object)
    Dim i As Long
    Dim objImgInf As clsBgImgInfo
    
    For i = 0 To UBound(objImgInfos)
        If Not objImgInfos(i) Is Nothing Then
            Set objImgInf = objImgInfos(i).CopyNew
            objImgInf.ImgCommand = icDownload
            objImgInf.LoadState = lsNone
            
            Call Me.ucBgImages.ConstructionImgData(objImgInf)
        End If
    Next
    
    Call Me.ucBgImages.Refresh
End Sub

Public Function ZlShowMe(objParent As Object, ByVal lngAdviceId As Long, _
    objSelImg As DicomImage, objImgInfos() As Object, _
    Optional lngWindowType As TImgProcessType = ptPreview, Optional lngPreviewTime As Long = 0, _
    Optional blnIsReportShow As Boolean) As Boolean
'lngType:窗口类型，0-图像处理窗口；1-图像预览窗口；2-标记图处理窗口
    
    On Error GoTo err
    
    Dim i As Integer
    Dim oldWinType As TImgProcessType
    Dim arrImages() As String
 
    oldWinType = mlngWinType
     
    mlngWinType = lngWindowType
    mlngPreViewTime = lngPreviewTime
    mblnIsReportShow = blnIsReportShow
    mstrUser = GetUserInfo
    
    If mblnIsChanged Then
        If MsgBoxD(Me, "有尚未保存的图像处理，该操作将清空这些处理，是否继续？", vbYesNo, Me.Caption) = vbNo Then
            Exit Function
        End If
    End If
    
    mblnDrag = False
    mblnIsChanged = False
    mblnCase = False
    
    Timer1.Enabled = False
    Timer2.Enabled = False
    
    If mlngWinType = ptPreview Then mlngState = 1
    If mlngWinType = ptProcess Then mlngState = 2
    
    If mlngWinType <> ptMark Then
        Me.ucBgImages.IsShowCheck = False
        
        '载入检查图像
        If mlngAdviceId <> lngAdviceId Then
            Call Me.ucBgImages.ClearAll
            
            '如果直接是图像处理，则同时加载缩略图
            If mlngWinType = ptProcess Then
                Call LoadImgs(objImgInfos)
            Else
                '先进行数组赋值，在timer1中延后加载
                maryImgInfos = objImgInfos
            End If
        Else
            '如果是先弹出预览窗口，然后在进行图像处理，则需要判断图像数量是否为0，因为在预览时，是没有加载缩略图的
            If mlngWinType = ptProcess And ucBgImages.ImgCount <= 0 Then
                Call LoadImgs(objImgInfos)
            End If
        End If
    Else
        Call Me.ucBgImages.ClearAll
    End If
    
    Set mOldImage = objSelImg
        
    Me.DViewer.Images.Clear
    Me.DViewer.Images.Add objSelImg
     
    If mlngAdviceId <> lngAdviceId Or Me.Visible = False Then
        Me.Show 0, objParent
        SetWindowPos Me.hwnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3  '将窗口置顶
    End If
    
    Call RefrshObjVisible
    
    
    If mlngWinType = ptPreview Then
        If DViewer.Images.Count > 0 Then
            Call DrawHintTag(DViewer.Images(1))
        End If
            
        Timer1.Enabled = True

        If lngPreviewTime > 0 Then
            Timer2.Interval = lngPreviewTime * 1000
            Timer2.Enabled = True
        End If
    Else
        refreshFace
    End If
    
    If oldWinType <> mlngWinType Then Call RestorceWinLayout
    
    mlngAdviceId = lngAdviceId
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
End Function



Private Sub RefrshObjVisible()
    Dim blnVisible As Boolean
    
    If mlngWinType = ptMark Then
        Me.cbrMain.FindControl(, conMenu_Process_Window).Parent.Visible = True
        Me.ucSplitter.Visible = False
        Me.ucBgImages.Visible = False
        Me.lblMemoText.Visible = False
        Me.cbxMemoText.Visible = False
        Me.picCboDropDown.Visible = False
        Me.cmdFont.Visible = False
        Me.cmdAdd.Visible = False
        Me.picMemo.Visible = False
        Me.lstMemoText.Visible = False
        Me.txtInputText.Visible = False
        
        Me.Caption = "标记图"
    Else
        If mlngState = 1 Then
            blnVisible = False
        Else
            blnVisible = True
        End If
        
        Me.ucBgImages.Visible = blnVisible
        Me.ucSplitter.Visible = blnVisible
        Me.lblMemoText.Visible = blnVisible
        Me.cbxMemoText.Visible = blnVisible
        Me.picCboDropDown.Visible = blnVisible
        Me.cmdInsert.Visible = blnVisible
        Me.cmdFont.Visible = blnVisible
        Me.cmdAdd.Visible = blnVisible
        Me.picMemo.Visible = blnVisible
        
        Me.cbrMain.FindControl(, conMenu_Process_Window).Parent.Visible = blnVisible

        If Me.lstMemoText.Visible Then Me.lstMemoText.Visible = blnVisible
        If Me.txtInputText.Visible Then Me.txtInputText.Visible = blnVisible
        
        Me.Caption = IIf(mlngWinType = ptPreview, "图像预览", "图像处理")
    End If
     
End Sub

Private Sub ClearLable(dcmImage As DicomImage)
    Dim i As Long
     '去除边框
    For i = 1 To dcmImage.Labels.Count
        If dcmImage.Labels(i).tag = "SELECT" Or dcmImage.Labels(i).tag = "BORDER" Or dcmImage.Labels(i).tag = "HINT" Then
            dcmImage.Labels(i).Visible = False
        End If
    Next
    dcmImage.BorderColour = vbWhite
End Sub

Private Function getListIndex() As Integer
'根据检索条件获取索引
    Dim i As Integer

    getListIndex = -1
    
    If mrsTmp.RecordCount <= 0 Then Exit Function

    mrsTmp.MoveFirst
    
    If cbxMemoText.Text = "" Then Exit Function

    For i = 0 To mrsTmp.RecordCount - 1
        If InStr(Trim(nvl(mrsTmp!简码)), UCase(cbxMemoText.Text)) > 0 Or InStr(Trim(nvl(mrsTmp!名称)), UCase(cbxMemoText.Text)) > 0 Then
            getListIndex = i
            
            Exit For
        End If

        mrsTmp.MoveNext
    Next
End Function

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Select Case Control.ID
        Case conMenu_Process_Save           '标记图保存
            If Control.Visible Then Control.Enabled = mblnAllowSaveReportImg
            
        Case conMenu_Process_SaveToStudy         '保存到检查
            If Control.Visible Then Control.Enabled = mblnAllowSaveStudyImg
            
        Case conMenu_Process_SaveToReport           '保存到报告
            If Control.Visible Then Control.Enabled = mblnAllowSaveReportImg
    End Select
Exit Sub
errHandle:
End Sub

Private Sub cbxMemoText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        If lstMemoText.ListCount > 0 Then lstMemoText.ListIndex = getListIndex
    End If
End Sub

Private Sub cmdAdd_Click()
'------------------------------------------------
'功能：添加操作
'参数：
'返回：无
'------------------------------------------------
    On Error GoTo err

    '拼接方法
    Call subAddMemoText


    Me.DViewer.Refresh

    '清空ComboBox文本
    cbxMemoText.Text = ""

    '关闭下拉框
    lstMemoText.Visible = False
     
    Call EnterProcessState
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
End Sub

Private Function GetNewImage(emImageType As TImageType) As DicomImage
    Dim dcmImage As DicomImage
    Dim img As New DicomImage
    Dim iPlane As Integer
    Dim aryDcm() As Byte
      
    If Me.DViewer.Images.Count = 1 Then
        Set dcmImage = Me.DViewer.Images(1)
          
        If emImageType <> mtTagImage Then
On Error GoTo errRead
            '转换一次图片格式，保存标注
            aryDcm = dcmImage.ArrayExport("BMP")
            img.ArrayImport aryDcm, "BMP"
errRead:
            If err.Number <> 0 Then
                '解决图片裁剪小了后报错问题
                Set GetNewImage = dcmImage.PrinterImage(8, iPlane, True, 1, 0, dcmImage.SizeX, 0, dcmImage.SizeY)
            Else
                Set GetNewImage = img
            End If
            
            err.Clear
            
            GetNewImage.InstanceUID = CreateUID
            GetNewImage.SeriesUID = dcmImage.SeriesUID
            GetNewImage.StudyUID = dcmImage.StudyUID
            
            If emImageType = mtReportImage Then
                GetNewImage.BorderWidth = 1
                GetNewImage.BorderColour = vbWhite
            End If
        Else
            dcmImage.InstanceUID = CreateUID
            
            Set GetNewImage = dcmImage
        End If
    Else
        Set GetNewImage = Nothing
    End If
End Function

Private Sub SaveImage(emImageType As TImageType)
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim objDcmImg As DicomImage
    Dim objSourceImgInfo As clsBgImgInfo
    
    Set objDcmImg = GetNewImage(emImageType)
    If objDcmImg Is Nothing Then Exit Sub
    
    If emImageType <> mtTagImage Then
        '如果为标记图处理时，是不存在对应缩略图显示的
        Set objSourceImgInfo = ucBgImages.ImageInfo(0).CopyNew()
        
        objSourceImgInfo.Key = objDcmImg.InstanceUID
        objSourceImgInfo.Filename = objDcmImg.InstanceUID
        objSourceImgInfo.ImgCommand = icReadly
        objSourceImgInfo.LoadState = lsNone
        objSourceImgInfo.Format = ifDcm
        objSourceImgInfo.JpgConvert = True
        objSourceImgInfo.IsReDrawed = False
        objSourceImgInfo.ErrorInfo = ""
        objSourceImgInfo.Redo = 0
        
        
        If FileExists(objSourceImgInfo.FilePath & objSourceImgInfo.Filename) = False Then
            objDcmImg.WriteFile objSourceImgInfo.FilePath & objSourceImgInfo.Filename, True, "1.2.840.10008.1.2.1"
        End If
        
        RaiseEvent OnSaveImage(emImageType, objDcmImg)
        
        strSQL = "select a.图像号,b.序列号 from 影像检查图象 a , 影像检查序列 b where a.图像UID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询图像号", objDcmImg.InstanceUID)
        
        If rsData.RecordCount > 0 Then
            objSourceImgInfo.ImageOrder = nvl(rsData!图像号)
            objSourceImgInfo.SeriesNoTag = nvl(rsData!序列号)
        End If
        
        ucBgImages.AddImg objSourceImgInfo
    Else
        RaiseEvent OnSaveImage(emImageType, objDcmImg)
    End If
    
    
    mblnIsChanged = False
 
    If emImageType = mtTagImage Then
        Unload Me
    End If
End Sub


Private Sub DViewer_Click()
On Error GoTo err
     
    Call EnterProcessState
        
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub DViewer_DblClick()
    Dim ls As DicomLabels
    Dim l As DicomLabel
    
    On Error GoTo err
    
    If mintMouseState = msMove Then
        Set ls = DViewer.LabelHits(mlngBaseXX, mlngBaseYY, False, False, True)
        If ls.Count > 0 Then
            If MsgBoxD(Me, "是否删除这个标注？", vbOKCancel, "提示") = vbOK Then
                Set l = ls(1)
                If l.tag <> "" Then
                    '是编号标注，需要同时删除三个标注，先删掉两个
                    If DViewer.Images(1).Labels.IndexOf(l.TagObject.TagObject) <> 0 Then
                        Call DViewer.Images(1).Labels.Remove(DViewer.Images(1).Labels.IndexOf(l.TagObject.TagObject))
                    End If
                    If DViewer.Images(1).Labels.IndexOf(l.TagObject) <> 0 Then
                        Call DViewer.Images(1).Labels.Remove(DViewer.Images(1).Labels.IndexOf(l.TagObject))
                    End If
                End If
                '是普通标注，或者编号的最后一个标注，直接删除即可
                Call DViewer.Images(1).Labels.Remove(DViewer.Images(1).Labels.IndexOf(l))
                DViewer.Refresh
            End If
        End If
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

'Private Sub DViewer_MouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)
'    Call ucMiniature.MouseWheel(Delta)
'End Sub

Private Function ThumbnailImgCount()
On Error GoTo errHandle
    ThumbnailImgCount = UBound(maryImgInfos()) + 1
Exit Function
errHandle:
    ThumbnailImgCount = 0
End Function
  
Private Sub Form_Terminate()
    Dim i As Long
    
    Set mImage = Nothing
    Set mdcmSelectLabel = Nothing
    Set mMovingLabel = Nothing
    Set mOldImage = Nothing
    Set mrsTmp = Nothing
    
    For i = 0 To ThumbnailImgCount - 1
        Set maryImgInfos(i) = Nothing
    Next
    
    Erase maryImgInfos
End Sub

Private Sub lstMemoText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call zlControl.CboSetText(cbxMemoText, lstMemoText.list(lstMemoText.ListIndex))
End Sub

Private Sub picCboDropDown_Click()
    lstMemoText.ZOrder
    lstMemoText.Visible = Not lstMemoText.Visible
    If lstMemoText.ListCount > 0 Then lstMemoText.ListIndex = getListIndex

    If lstMemoText.Visible Then lstMemoText.SetFocus
     
    Call EnterProcessState
End Sub

Private Sub EnterProcessState()
    mlngState = 3
    
    mlngPreViewTime = 0
    
    Timer1.Enabled = False
    Timer2.Enabled = False
     
    If mlngWinType = ptPreview Then mlngWinType = ptProcess
End Sub

Private Sub cbxMemoText_Change()
    lstMemoText.ZOrder
    
    If Not lstMemoText.Visible Then lstMemoText.Visible = True
    If lstMemoText.ListCount > 0 Then lstMemoText.ListIndex = getListIndex
End Sub

Private Sub cbxMemoText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cbxMemoText.ListIndex = lstMemoText.ListIndex
        lstMemoText.Visible = False
        
        cbxMemoText.SelStart = 0
        cbxMemoText.SelLength = Len(cbxMemoText.Text)
        cbxMemoText.SetFocus
    End If
    
    If KeyAscii = vbKeyEscape Then lstMemoText.Visible = False
End Sub

Private Sub cmdFont_Click()
On Error GoTo errHandle
    diaFont.flags = 1
    diaFont.FontBold = Me.Font.Bold
    diaFont.FontItalic = Me.Font.Italic
    diaFont.FontName = Me.Font.Name
    diaFont.FontSize = Me.Font.Size
    diaFont.FontStrikethru = Me.Font.Strikethrough
    diaFont.FontUnderline = Me.Font.Underline

    
    diaFont.ShowFont
    
    Me.Font.Bold = diaFont.FontBold
    Me.Font.Italic = diaFont.FontItalic
    Me.Font.Name = diaFont.FontName
    Me.Font.Size = diaFont.FontSize
    Me.Font.Strikethrough = diaFont.FontStrikethru
    Me.Font.Underline = diaFont.FontUnderline
    
    Call EnterProcessState
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngCount As Long
    
    On Error GoTo errHandle
     
    Call EnterProcessState
    
    mblnDrag = False
    Select Case Control.ID
        Case conMenu_Process_Window         '亮度对比度
            subSetMouseState 1
            'Control.Checked = True
            
        Case conMenu_Process_Zoom           '缩放
            subSetMouseState 2
            'Control.Checked = True
            
        Case conMenu_Process_RectZoom       '裁剪缩放
            subSetMouseState msRectangle
            'Control.Checked = True
        
        Case conMenu_Process_RectCapture         '裁剪后采集
            Call CaptureFrameSelectImage
            
        Case conMenu_Process_RRotate        '顺时针旋转
            subSetRotate True
            
        Case conMenu_Process_LRotate        '逆时针旋转
            subSetRotate False
            
        Case conMenu_Process_Sharpness      '锐化
            subSetSharp True
            
        Case conMenu_Process_Filter         '平滑
            subSetSharp False
        
        Case conMenu_Process_Line           '直线标注
            subSetMouseState msline
        
        Case conMenu_Process_Arrow          '箭头标注
            subSetMouseState msArrow
            
        Case conMenu_Process_Ellipse        '圆形标注
            subSetMouseState msEllipse
            
        Case conMenu_Process_TextTag           '文字标注
            mstrText = ""
            subSetMouseState msText
            
        Case conMenu_Process_DelAllLabels   '清除标注
            lngCount = DViewer.Images(1).Labels.Count
            
            DViewer.Images(1).Labels.Clear
            DViewer.Refresh
            
            mintAutoNumber = 0
            
            If lngCount <> DViewer.Images(1).Labels.Count Then
                mblnIsChanged = True
            End If
            
        Case conMenu_Process_MoveLabel      '移动
            mblnDrag = True
            subSetMouseState msMove
            
        Case conMenu_Process_LabelSetUp     '标注设置
            Call subSetTextLabel
            
        Case conMenu_Process_Restore        '恢复
            DViewer.Images.Clear
            DViewer.Images.Add mOldImage
            
            If DViewer.Images.Count > 0 Then
                ClearLable DViewer.Images(1)
            End If
            
            '重建标注之间的关联
            Call subLabelCopyRebuild(mOldImage, Me.DViewer.Images(1))
            mintAutoNumber = 0  '恢复打开图像时的最大序号
            
        Case conMenu_Process_Num * 100 To conMenu_Process_Num * 100 + 9
            mintNumberIndex = Val(Control.Category)
            subSetMouseState msNumber
        
        Case conMenu_Process_NumTag
            mintNumberIndex = 0
            subSetMouseState msNumber
            
        Case conMenu_Process_Word * 100 To conMenu_Process_Word * 100 + 99
            mstrText = Control.Caption
            subSetMouseState msFixText
        
        Case conMenu_Process_Save           '标记图保存
            If mlngWinType = ptMark Then
                Call SaveImage(mtTagImage)
            End If
            
        Case conMenu_Process_SaveToStudy         '保存到检查
            Call SaveImage(mtStudyImage)
            
            
        Case conMenu_Process_SaveToReport           '保存到报告
            Call SaveImage(mtReportImage)
            
        Case conMenu_Process_Exit             '退出
            Unload Me
    End Select
    
'    If Control.ID <> conMenu_Process_LabelSetUp Then
'        Call setCmdLabelColor
'    End If
    If mlngWinType = ptPreview Then
        Me.Caption = "图像处理"
        mlngWinType = ptProcess
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub



Private Sub subSetSharp(blnSharp As Boolean)
'------------------------------------------------
'功能：dcmView中图像的平滑和锐化
'参数：blnSharp表示图像处理的方向，True=锐化；False=平滑
'返回：无，直接处理dcmView中的图像
'------------------------------------------------
    If DViewer.Images.Count > 0 Then
        If blnSharp = True Then
            '锐化处理
            If DViewer.Images(1).FilterLength <= 0 Then
                DViewer.Images(1).FilterLength = 0
                '先前没有平滑处理，直接进行锐化处理
                DViewer.Images(1).UnsharpEnhancement = DViewer.Images(1).UnsharpEnhancement + 0.1
            Else
                '如果先前已经有平滑处理，则先淡化平滑效果
                DViewer.Images(1).FilterLength = DViewer.Images(1).FilterLength - 1
            End If
        Else
            '平滑处理
            '判断Zoom是否＝1，如果是，则修改为0.9999
            If DViewer.Images(1).ActualZoom = 1 Then
                DViewer.Images(1).Zoom = 0.9999
            End If
            
            If DViewer.Images(1).UnsharpEnhancement <= 0 Then
                DViewer.Images(1).UnsharpEnhancement = 0
                '先前没有锐化处理，直接开始平滑
                '判断FilterLength是否＝0如果是，则在2/ActualZoom和2×FilterLength之间进行调整
                If DViewer.Images(1).FilterLength = 0 Then
                    DViewer.Images(1).FilterLength = 2 / DViewer.Images(1).ActualZoom + 1
                Else    '正常情况下FilterLength＋1
                    DViewer.Images(1).FilterLength = DViewer.Images(1).FilterLength + 1
                End If
            Else
                '先前已经有了锐化处理，先淡化锐化的效果
                DViewer.Images(1).UnsharpEnhancement = DViewer.Images(1).UnsharpEnhancement - 0.1
            End If
        End If
    End If
    
    mblnIsChanged = True
End Sub

Private Sub subSetRotate(blnClockwise As Boolean)
'------------------------------------------------
'功能：dcmView中图像的旋转
'参数：blnClockwise旋转的方向,True=顺时针旋转；False=逆时针旋转
'返回：无，直接处理dcmView中的图像
'------------------------------------------------
    If DViewer.Images.Count > 0 Then
        Dim iRotateState As Integer
        
        iRotateState = DViewer.Images(1).RotateState
        If blnClockwise = True Then
            iRotateState = iRotateState - 1
        Else
            iRotateState = iRotateState + 1
        End If
        If iRotateState = -1 Then iRotateState = 3
        iRotateState = iRotateState Mod 4
        DViewer.Images(1).RotateState = iRotateState
    End If
End Sub


'DicomViewer裁剪后采集图象
Private Sub CaptureFrameSelectImage()
    Dim imgResult As DicomImage
    Dim imgs As New DicomImages
    Dim iPlane As Integer
    Dim dblZoom As Double
    Dim iLeft As Integer
    Dim iRight As Integer
    Dim iTop As Integer
    Dim iBottom As Integer
    Dim iMax As Integer
    Dim img As DicomImage
    Dim lblFrame As DicomLabel
    
    If Me.DViewer.Images.Count <> 1 Then Exit Sub
    If Me.DViewer.Images(1).Labels.Count < 1 Then Exit Sub
    
    Set img = Me.DViewer.Images(1)
    Set lblFrame = Me.DViewer.Images(1).Labels(Me.DViewer.Images(1).Labels.Count)
    
    If Abs(lblFrame.Width) = 0 Or Abs(lblFrame.Height) = 0 Then
        MsgBoxD Me, "请选择图像区域后再保存", vbExclamation, "提示"
        Exit Sub
    End If
    
    '图象最大宽高=300
    iMax = 300
    
    '根据label来提取被框选中的图像
    '图象位数,黑白图像为1，彩色图像为3
    iPlane = 1
    If Not IsNull(img.Attributes(&H28, &H4).value) And img.Attributes(&H28, &H4).Exists Then
        If img.Attributes(&H28, &H4).value = "RGB" Or img.Attributes(&H28, &H4).value = "YBR_FULL_422" Then
            iPlane = 3
        End If
    End If
    
    '图象框的位置
    If lblFrame.Width >= 0 Then
        iLeft = lblFrame.Left
        iRight = iLeft + lblFrame.Width
    Else
        iLeft = lblFrame.Left + lblFrame.Width
        iRight = lblFrame.Left
    End If
    
    If lblFrame.Height >= 0 Then
        iTop = lblFrame.Top
        iBottom = iTop + lblFrame.Height
    Else
        iTop = lblFrame.Top + lblFrame.Height
        iBottom = lblFrame.Top
    End If
    
    '控制结果图象的大小在300*300之内
    If (iRight - iLeft) > iMax Or (iBottom - iTop) > iMax Then
        dblZoom = iMax / (iRight - iLeft)
        If dblZoom > iMax / (iBottom - iTop) Then dblZoom = iMax / (iBottom - iTop)
    Else
        dblZoom = 1
    End If
    
    img.Labels(img.Labels.Count).Visible = False
    If (img.RotateState = doRotateLeft And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipBoth) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipHorizontal) Then
        'X方向对调
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, img.SizeX - iRight, img.SizeX - iLeft, iTop, iBottom)
    ElseIf (img.RotateState = doRotateLeft And img.FlipState = doFlipBoth) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipVertical) Then
        'Y方向对调
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, img.SizeY - iBottom, img.SizeY - iTop)
    ElseIf (img.RotateState = doRotateRight And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateLeft And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipBoth) Then
        'X，Y方向对调
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, img.SizeX - iRight, img.SizeX - iLeft, img.SizeY - iBottom, img.SizeY - iTop)
    Else
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, iTop, iBottom)
    End If
    
    DViewer.Images.Clear
    DViewer.Images.Add imgResult
    
    mblnCase = True
    mblnIsChanged = True
End Sub

Private Sub subSetMouseState(intMoustState As TMouseState)
'------------------------------------------------
'功能：设置鼠标状态，同时更新工具栏按钮的选择状态
'参数：intMoustState -- 鼠标状态
'返回：无
'------------------------------------------------
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Window).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Zoom).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_RectZoom).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Arrow).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Ellipse).Checked = False
    cbrMain.FindControl(xtpControlSplitButtonPopup, conMenu_Process_TextTag).Checked = False
    cbrMain.FindControl(xtpControlSplitButtonPopup, conMenu_Process_NumTag).Checked = False
'    cbrMain.FindControl(xtpControlButton, conMenu_Process_Corp).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_MoveLabel).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Line).Checked = False
     
    '改变当前鼠标状态
    If mintMouseState = intMoustState Then
        If intMoustState = msNumber Or intMoustState = msFixText Then
            If mintNumberIndex > 0 Or Len(mstrText) > 0 Then
                mintMouseState = intMoustState
                Exit Sub
            End If
        End If
        
        mintMouseState = msNone
    Else
        mintMouseState = intMoustState
        
        Select Case mintMouseState
            Case msWinLevel: cbrMain.FindControl(xtpControlButton, conMenu_Process_Window).Checked = True
            Case msZoom: cbrMain.FindControl(xtpControlButton, conMenu_Process_Zoom).Checked = True
            Case msRectangle: cbrMain.FindControl(xtpControlButton, conMenu_Process_RectZoom).Checked = True
            Case msArrow: cbrMain.FindControl(xtpControlButton, conMenu_Process_Arrow).Checked = True
            Case msEllipse: cbrMain.FindControl(xtpControlButton, conMenu_Process_Ellipse).Checked = True
            Case msText
                If mstrText = "" Then
                    cbrMain.FindControl(xtpControlSplitButtonPopup, conMenu_Process_TextTag).Checked = True
                End If
            Case msNumber
                If mintNumberIndex = 0 Then
                    cbrMain.FindControl(xtpControlSplitButtonPopup, conMenu_Process_NumTag).Checked = True
                End If
            Case msMove: cbrMain.FindControl(xtpControlButton, conMenu_Process_MoveLabel).Checked = True
            Case msline: cbrMain.FindControl(xtpControlButton, conMenu_Process_Line).Checked = True
        End Select
    End If
    
End Sub

Private Sub cbrMain_Resize()
    Call refreshFace
     
End Sub

Public Sub refreshFace()
    '设置显示的客户区域
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    On Error Resume Next
    
    cbrMain.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    
     
    If mlngWinType = ptMark Then
        Me.DViewer.Left = lngLeft
        Me.DViewer.Top = lngTop
        Me.DViewer.Width = lngRight
        Me.DViewer.Height = lngBottom - lngTop
    Else
        If mlngState <= 1 Then
            Me.DViewer.Left = lngLeft
            Me.DViewer.Top = lngTop
            Me.DViewer.Width = lngRight
            Me.DViewer.Height = lngBottom - lngTop
        Else
            Me.ucSplitter.Top = lngTop
            Me.ucSplitter.Height = lngBottom - lngTop - 600
            
            Me.ucBgImages.Left = lngLeft
            Me.ucBgImages.Top = lngTop
            Me.ucBgImages.Height = lngBottom - lngTop - 600
    
            '摆放DViewer
            Me.DViewer.Left = lngLeft + Me.ucBgImages.Width + Me.ucSplitter.Width
            Me.DViewer.Top = lngTop
            Me.DViewer.Width = lngRight - lngLeft - Me.ucBgImages.Width - ucSplitter.Width
            Me.DViewer.Height = lngBottom - lngTop - 600
            
            ucSplitter.RePaint
            
            If lstMemoText.Visible Then lstMemoText.ZOrder
        End If
    End If
        
    Me.picMemo.Left = lngLeft
    Me.picMemo.Top = Me.DViewer.Top + Me.DViewer.Height
    Me.picMemo.Height = 600
    Me.picMemo.Width = lngRight
    
    Me.lstMemoText.Left = Me.cbxMemoText.Left
    Me.lstMemoText.Top = Me.picMemo.Top + Me.cbxMemoText.Top - Me.lstMemoText.Height
    Me.lstMemoText.Width = Me.cbxMemoText.Width - 10
    
    err.Clear
End Sub

Private Sub cmdInsert_Click()
    Dim strSQL As String, i As Integer
    Dim strUser As String
    
    Call EnterProcessState
    
    If Trim(cbxMemoText.Text) = "" Then
        MsgBoxD Me, "请输入备注内容。", vbInformation, "提示"
        If cbxMemoText.Enabled Then cbxMemoText.SetFocus
        Exit Sub
    End If
    
    If cbxMemoText.ListIndex <> -1 Then
        MsgBoxD Me, "该备注内容已经在常用备注中。", vbInformation, "提示"
        If cbxMemoText.Enabled Then cbxMemoText.SetFocus
        Exit Sub
    Else
        For i = 0 To cbxMemoText.ListCount - 1
            If UCase(Trim(cbxMemoText.list(i))) = UCase(Trim(cbxMemoText.Text)) Then
                MsgBoxD Me, "该备注容已经在常用备注中。", vbInformation, "提示"
                If cbxMemoText.Enabled Then cbxMemoText.SetFocus
                Exit Sub
            End If
        Next
    End If
        
    On Error GoTo errH
    
    strSQL = zlCommFun.zlGetSymbol(cbxMemoText.Text)
    strSQL = "zl_影像图像备注_Insert('" & Replace(cbxMemoText.Text, "'", "''") & "','" & strSQL & "','" & mstrUser & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    AddComboItem cbxMemoText.hwnd, CB_ADDSTRING, 0, cbxMemoText.Text
    lstMemoText.AddItem cbxMemoText.Text
    
    MsgBoxD Me, "已设置为常用备注。", vbInformation, "提示"
    
    If cbxMemoText.Enabled Then cbxMemoText.SetFocus
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub subAddMemoText()
'------------------------------------------------
'功能：给图像添加备注文字
'参数：
'返回：无
'------------------------------------------------
    On Error GoTo err
    
    Dim img As DicomImage
    Dim iLeft As Integer
    Dim iWidth As Integer
    Dim iTop As Integer
    Dim iHeight As Integer
    Dim imgResult As New DicomImage
    Dim iPlane As Integer
    Dim lngWhiteX As Long
    Dim lngWhiteY As Long
    Dim lngFontHeight As Long
    
    If Me.DViewer.Images.Count <> 1 Then Exit Sub
    
    If Trim(cbxMemoText.Text) = "" Then Exit Sub
    
    lngFontHeight = ScaleY(TextHeight(cbxMemoText.Text), vbTwips, vbPixels) + 6
    
    '把备注文字添加到图像中
    Set img = Me.DViewer.Images(1)
    
    iLeft = 0
    iTop = 0
    iWidth = img.SizeX
    iHeight = img.SizeY + lngFontHeight

    '使用PrinterImage方法，可以将图像上的标签及标注同时进行绘制
    Set imgResult = img.PrinterImage(8, iPlane, True, 1, 0, iWidth, 0, iHeight - lngFontHeight)
'

    '添加标注
    Dim dlMemoText As New DicomLabel
    
    dlMemoText.LabelType = doLabelText
    dlMemoText.ImageTied = True
    dlMemoText.Transparent = False
    dlMemoText.AutoSize = False
    dlMemoText.Left = 0
    dlMemoText.Top = img.SizeY
    dlMemoText.Width = iWidth
    dlMemoText.Height = lngFontHeight
    
    dlMemoText.BackColour = vbWhite
    dlMemoText.ForeColour = vbBlack
            
    dlMemoText.Font.Name = Me.Font.Name
    dlMemoText.Font.Italic = Me.Font.Italic
    dlMemoText.Font.Strikethrough = Me.Font.Strikethrough
    dlMemoText.Font.Underline = Me.Font.Underline
    dlMemoText.Font.Size = Me.Font.Size
    dlMemoText.Font.Bold = Me.Font.Bold
    dlMemoText.FontName = Me.Font.Name
    dlMemoText.FontSize = Me.Font.Size
    dlMemoText.ShowTextBox = True
    
    dlMemoText.Text = Me.cbxMemoText.Text & "                                                                                                                                 "
    
    imgResult.Labels.Add dlMemoText
    
    Set imgResult = imgResult.PrinterImage(8, iPlane, True, 1, 0, iWidth, 0, iHeight)

    '更新图像
    Me.DViewer.Images.Clear
    Me.DViewer.Images.Add imgResult
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
End Sub

Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim ls As DicomLabels
    Dim lngLeftD As Long
    
    If mlngWinType = ptPreview Then Exit Sub
    
    If Button = 1 And DViewer.Images.Count > 0 Then
        Dim intLabelType As Integer
        
        If mblnDrag Then
            Set ls = DViewer.LabelHits(X, Y, False, False, True)
            mlngBaseX = DViewer.ImageXPosition(X, Y)
            mlngBaseY = DViewer.ImageYPosition(X, Y)
            If ls.Count > 0 Then    '如果选中了任何一个标注
                '如果Tag=""说明是简单标注，非空说明是数字编号标注，需要找到文字标注
                mintMouseState = msMove
                Set mMovingLabel = ls(1)
                If mMovingLabel.tag <> "" Then
                    If mMovingLabel.tag = m_LabelTag_Back Then
                        Set mMovingLabel = mMovingLabel.TagObject
                    ElseIf mMovingLabel.tag = m_LabelTag_Circle Then
                        Set mMovingLabel = mMovingLabel.TagObject.TagObject
                    End If
                End If
            Else
                mintMouseState = msDrag
            End If
        End If
                    
        mMouseDownPoint.X = DViewer.Images(1).ActualScrollX
        mMouseDownPoint.Y = DViewer.Images(1).ActualScrollY
          
        mInitScrollPoint.X = DViewer.Images(1).ScrollX + X
        mInitScrollPoint.Y = DViewer.Images(1).ScrollY + Y
        
        mblnDcmViewDown = True
        If mintMouseState <> msNone Then
            '记录当前鼠标位置
            mlngBaseXX = X
            mlngBaseYY = Y
            Select Case mintMouseState
                Case msline, msArrow, msEllipse, msText, msRectangle, msFixText, msNumber     '直线，箭头，椭圆，文字，框选，固定文字，顺序编号
                    If mintMouseState = msArrow Then
                        intLabelType = doLabelArrow
                    ElseIf mintMouseState = msEllipse Or mintMouseState = msNumber Then
                        intLabelType = doLabelEllipse
                    ElseIf mintMouseState = msText Or mintMouseState = msFixText Then
                        intLabelType = doLabelText
                    ElseIf mintMouseState = msRectangle Then
                        intLabelType = doLabelRectangle
                    ElseIf mintMouseState = msline Then
                        intLabelType = doLabelLine
                    End If
                    
                    If mintMouseState = msFixText Then
                        '如果是单个文字，位移的量要减少
                        If mstrText = "自定义" Then
                            lngLeftD = IIf(Len(mstrCustom) = 1, 3, 7)
                        Else
                            lngLeftD = IIf(Len(Left(mstrText, InStr(mstrText & "=", "=") - 1)) = 1, 3, 7)
                        End If
                    Else
                        lngLeftD = 7
                    End If
                    DViewer.Images(1).Labels.Add GetNewLabel(intLabelType, DViewer.ImageXPosition(X, Y) - lngLeftD, DViewer.ImageYPosition(X, Y) - 7, 0, 0)
                    Set mdcmSelectLabel = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count)
                    If intLabelType = doLabelArrow Then
                        '箭头需要使用线宽=2
                        mdcmSelectLabel.LineWidth = 4
                    ElseIf intLabelType = doLabelLine Then
                        mdcmSelectLabel.LineWidth = 2
                    ElseIf intLabelType = doLabelText Then
                        mdcmSelectLabel.XOR = False
                        mdcmSelectLabel.ForeColour = vbBlack
                        If mlngWinType <> ptMark Then
                            '不是标记图，则给文字增加背景，标记图不能增加，因为电子病历不支持，打印的时候就不支持
                            mdcmSelectLabel.Transparent = False
                            mdcmSelectLabel.ForeColour = vbWhite
                            mdcmSelectLabel.BackColour = vbBlack
                        End If
                        '设置字体大小
                        If DViewer.Images(1).SizeX <= 256 Then
                            mdcmSelectLabel.FontSize = 10
                        ElseIf DViewer.Images(1).SizeX <= 512 Then
                            mdcmSelectLabel.FontSize = 15
                        Else
                            mdcmSelectLabel.FontSize = 18
                        End If
                        
                    End If
                    
                    mblnIsChanged = True
            End Select
        End If
    End If
End Sub

Private Sub DViewer_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    
    If mlngWinType = ptPreview Then Exit Sub
    
    If mblnDcmViewDown = True And Button = 1 And DViewer.Images.Count > 0 Then
        Select Case mintMouseState
            Case msWinLevel   '亮度对比度
                DViewer.Images(1).Width = DViewer.Images(1).Width + (X - mlngBaseXX)
                DViewer.Images(1).Level = DViewer.Images(1).Level + (Y - mlngBaseYY)
                mlngBaseXX = X
                mlngBaseYY = Y
                mblnIsChanged = True
            Case msZoom   '缩放
                Dim dblZoom As Double
                dblZoom = DViewer.Images(1).ActualZoom
                dblZoom = dblZoom * (1 + (Y - mlngBaseYY) * 0.001)
                If dblZoom < 64 And dblZoom > 0.01 Then
                    subCenterZoom DViewer.Images(1), DViewer, dblZoom, mCorpSize
                    mblnIsChanged = True
                End If
                mlngBaseYY = Y
'            Case msRectangle  '裁剪缩放
'                Dim dcmLabel As DicomLabel
'                dcmView.Labels.Clear
'                Set dcmLabel = dcmView.Labels.AddNew
'                dcmLabel.LabelType = doLabelRectangle
'                dcmLabel.Left = mlngBaseXX
'                dcmLabel.Top = mlngBaseYY
'                dcmLabel.Width = x - mlngBaseXX
'                dcmLabel.Height = y - mlngBaseYY
            Case msline, msArrow, msEllipse, msRectangle    '直线,箭头标注'圆形标注,框选
                mdcmSelectLabel.Width = DViewer.ImageXPosition(X, Y) - mdcmSelectLabel.Left
                mdcmSelectLabel.Height = DViewer.ImageYPosition(X, Y) - mdcmSelectLabel.Top
                
                mblnIsChanged = True
            Case msDrag
                '拖动图像......
                DViewer.Images(1).ScrollX = mInitScrollPoint.X - X
                DViewer.Images(1).ScrollY = mInitScrollPoint.Y - Y
                
                mblnIsChanged = True
            Case msMove
                '移动标注
                If Not mMovingLabel Is Nothing Then
                    subaCorrectCursor DViewer, DViewer.Images(1), X, Y  '鼠标移动如果超出图像范围，则修正鼠标位置
                    subMoveLable mMovingLabel, DViewer.ImageXPosition(X, Y) - mlngBaseX, DViewer.ImageYPosition(X, Y) - mlngBaseY
                    mlngBaseX = DViewer.ImageXPosition(X, Y)
                    mlngBaseY = DViewer.ImageYPosition(X, Y)
                    
                    mblnIsChanged = True
                End If
        End Select
        
        DViewer.Refresh
    End If
End Sub

Private Sub DViewer_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mlngWinType = ptPreview Then Exit Sub

    If mblnDcmViewDown = True And Button = 1 And DViewer.Images.Count > 0 Then
        mblnDcmViewDown = False
        If mintMouseState = msText Then      '文字标注
            
            txtInputText.Left = Me.ScaleX(X, vbPixels, vbTwips) + DViewer.Left
            txtInputText.Top = Me.ScaleY(Y, vbPixels, vbTwips) + DViewer.Top
            
            txtInputText.Text = ""
            txtInputText.Visible = True
            txtInputText.SetFocus
            mblnIsChanged = True
        ElseIf mintMouseState = msRectangle Then   '裁剪缩放
            '显示图像保存菜单
            Call ShowFrameSelectImagePopup
            
            '删除框选用的临时标注
            If DViewer.Images(1).Labels.Count > 0 Then
                DViewer.Images(1).Labels.Remove DViewer.Images(1).Labels.Count
            End If
            
            Set mdcmSelectLabel = Nothing
            
'            dcmView.Labels.Clear
'            RectangleZoom dcmView, dcmView.Images(1), mlngBaseXX, mlngBaseYY, x - mlngBaseXX, y - mlngBaseYY
        ElseIf mintMouseState = msDrag Then
            '计算图像漫游的偏移位置
            mCorpSize.X = mCorpSize.X + (DViewer.Images(1).ActualScrollX - mMouseDownPoint.X)
            mCorpSize.Y = mCorpSize.Y + (DViewer.Images(1).ActualScrollY - mMouseDownPoint.Y)
            
            mblnIsChanged = True
        ElseIf mintMouseState = msFixText Then
            '添加固定文字
            If mstrText = "自定义" Then   '自定义文字标注
                mdcmSelectLabel.Text = mstrCustom
            Else
                mdcmSelectLabel.Text = Left(mstrText, InStr(mstrText & "=", "=") - 1)
            End If
            mblnIsChanged = True
        ElseIf mintMouseState = msNumber Then
            Dim intText As Integer
            
            If mintNumberIndex = 0 Then '自动顺序编号
                mintAutoNumber = mintAutoNumber + 1
                intText = mintAutoNumber
            Else
                intText = mintNumberIndex
            End If
            '添加顺序编号
            mdcmSelectLabel.XOR = False
            mdcmSelectLabel.BackColour = glngColor(intText Mod 9 + 1)
            mdcmSelectLabel.Transparent = False
            mdcmSelectLabel.Width = 14
            mdcmSelectLabel.Height = 14
            mdcmSelectLabel.tag = m_LabelTag_Back
            
            '添加顺序编号圆形的两个附加标注，圆形框和数字
            DViewer.Images(1).Labels.Add GetNewLabel(doLabelEllipse, mdcmSelectLabel.Left, mdcmSelectLabel.Top, 14, 14)
            Set mdcmSelectLabel = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count)
            mdcmSelectLabel.XOR = False
            mdcmSelectLabel.ForeColour = vbBlack
            mdcmSelectLabel.Transparent = True
            mdcmSelectLabel.tag = m_LabelTag_Circle
            mdcmSelectLabel.TagObject = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count - 1)
            
            DViewer.Images(1).Labels.Add GetNewLabel(doLabelText, mdcmSelectLabel.Left + 1, mdcmSelectLabel.Top, 0, 0)
            Set mdcmSelectLabel = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count)
            mdcmSelectLabel.ForeColour = vbBlack
            mdcmSelectLabel.Transparent = True
            mdcmSelectLabel.XOR = False
            mdcmSelectLabel.tag = m_LabelTag_Number
            mdcmSelectLabel.FontSize = 8
            mdcmSelectLabel.FontName = "Arial Bold"
            mdcmSelectLabel.AutoSize = True
            mdcmSelectLabel.Text = intText
            If mdcmSelectLabel.Text < 10 Then
                mdcmSelectLabel.Left = mdcmSelectLabel.Left + 3
            End If
            mdcmSelectLabel.TagObject = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count - 1)
            DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count - 2).TagObject = mdcmSelectLabel    'TagObject形成闭环
            
            mblnIsChanged = True
        End If
        
        DViewer.Refresh
    End If
End Sub

Public Sub ShowFrameSelectImagePopup()
'------------------------------------------------
'功能：创建框选图象的时候 ，鼠标右键的弹出菜单
'参数：
'返回：无
'------------------------------------------------

Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim cbrToolPopup As CommandBarPopup
    
    
    '鼠标右键弹出菜单
    Set cbrToolBar = Me.cbrMain.Add("鼠标右键", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RectCapture, "确认裁剪")
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub


Private Sub subCenterZoom(img As DicomImage, Viewer As DicomViewer, dblZoom As Double, corpSize As TPoint)
'------------------------------------------------
'功能：对图像进行缩放。以当前viewer中心点为缩放中心点。
'参数：
'       img -- 进行缩放的图像
'       viewer －－ 图像所在的viewer
'       dblZoom －－图像新的缩放倍数
'返回：无，直接调整图像的缩放倍数
'上级函数或过程：frmViewer.Viewer_MouseMove
'下级函数或过程：无
'引用的外部参数：无
'编制人： 黄捷 2006-2-10
'------------------------------------------------
    img.Zoom = dblZoom
    img.StretchToFit = False

            
    img.ScrollX = (img.SizeX * img.ActualZoom - ScaleX(Viewer.Width, vbTwips, vbPixels) / Viewer.MultiColumns) / 2 + corpSize.X
    img.ScrollY = (img.SizeY * img.ActualZoom - ScaleY(Viewer.Height, vbTwips, vbPixels) / Viewer.MultiRows) / 2 + corpSize.Y
End Sub


Private Sub Form_Load()
    On Error GoTo err

    '恢复窗体位置
    Call RestorceWinLayout
    
    '设置默认颜色
    glngColor(1) = RGB(186, 186, 186)
    glngColor(2) = RGB(255, 215, 0)
    glngColor(3) = RGB(255, 0, 255)
    glngColor(4) = RGB(255, 0, 130)
    glngColor(5) = RGB(0, 255, 0)
    glngColor(6) = RGB(130, 255, 255)
    glngColor(7) = RGB(255, 255, 0)
    glngColor(8) = RGB(0, 0, 255)
    glngColor(9) = RGB(0, 160, 0)
    
    Call subLoadTextLabel
    
    '创建工具栏
    Call InitCommandBars
    
    Call LoadMemoFontStyle
    
    mCorpSize.X = 0
    mCorpSize.Y = 0
    mblnOk = False
    mintAutoNumber = 0
    
    '图像处理，鼠标左键默认是调窗；图像标注，鼠标左键默认是移动标注
'    If mblnIsMark = True Then
'        Call subSetMouseState(msMove)
'    Else
'        Call subSetMouseState(msWinLevel)
'    End If
'
    ucBgImages.IsDrawOrder = False
    ucBgImages.IsDrawHint = False

    Call ReadEnjoin
    
    Call refreshFace
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

'载入备注字体样式
Private Sub LoadMemoFontStyle()
    Dim strFontStyle As String
    Dim aryFontStyle() As String
    
    '“宋体,12,B,U,S,I”
    
    strFontStyle = zlDatabase.GetPara("图像备注字体", glngSys, glngModul, "")
    
    strFontStyle = strFontStyle & ",,,,,,"
    
    aryFontStyle = Split(strFontStyle, ",")
    
    If aryFontStyle(0) <> "" Then Me.Font.Name = aryFontStyle(0)
    If Val(aryFontStyle(1)) <> 0 Then Me.Font.Size = Val(aryFontStyle(1))
    If UCase(aryFontStyle(2)) = "B" Then Me.Font.Bold = True
    If UCase(aryFontStyle(3)) = "U" Then Me.Font.Underline = True
    If UCase(aryFontStyle(4)) = "S" Then Me.Font.Strikethrough = True
    If UCase(aryFontStyle(5)) = "I" Then Me.Font.Italic = True
    
End Sub


Private Sub SaveMemoFontStyle()
    Dim strFontStyle As String
    
    strFontStyle = Me.Font.Name & "," & _
        Me.Font.Size & "," & _
        IIf(Me.Font.Bold, "B", "") & "," & _
        IIf(Me.Font.Underline, "U", "") & "," & _
        IIf(Me.Font.Strikethrough, "S", "") & "," & _
        IIf(Me.Font.Italic, "I", "")

    Call zlDatabase.SetPara("图像备注字体", strFontStyle, glngSys, glngModul)
End Sub


Private Function ReadEnjoin() As Boolean
'功能：读取并加入常用备注
    Dim strSQL As String, strPre As String
    Dim strUser As String
    
    On Error GoTo errH
    
    '常用嘱托
    strPre = cbxMemoText.Text '加入后保持原有值
    cbxMemoText.Clear
    
    strSQL = _
        " Select 名称,简码 From 影像图像备注 Where 名称 is Not Null And 人员=[1]" & _
        " Union" & _
        " Select 名称,简码 From 影像图像备注 Where 名称 is Not Null And 人员 is Null" & _
        " Order by 名称"
    Set mrsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUser)
    Do While Not mrsTmp.EOF
        AddComboItem cbxMemoText.hwnd, CB_ADDSTRING, 0, mrsTmp!名称
        
        lstMemoText.AddItem mrsTmp!名称
        mrsTmp.MoveNext
    Loop
    cbxMemoText.Text = strPre
    
    ReadEnjoin = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Form_Unload(Cancel As Integer)
    If mblnIsChanged Then
        If MsgBoxD(Me, "图像处理尚未保存，是否保存？", vbYesNo, "提示") = vbYes Then
            If mlngWinType = ptMark Then
                Call SaveImage(mtTagImage)
            Else
                Call SaveImage(mtStudyImage)
            End If
        End If
    End If
    
    mlngAdviceId = 0
    
    Call SaveWinLayout
    
    Call SaveMemoFontStyle
    
    RaiseEvent OnUnload
End Sub

Private Sub SaveWinLayout()
'保存窗体位置及界面布局
'由于默认窗口大小原因未使用ZL9COMLIB中的方法
    Dim strCaption As String
    Dim strPrivateReg As String
    
    strCaption = GetWindowCaption
    
    strPrivateReg = GetPrivateRegPath(strCaption)
    
    Call SaveSetting("ZLSOFT", strPrivateReg, "WinLeft", IIf(Me.Left < 0, 0, Me.Left))
    Call SaveSetting("ZLSOFT", strPrivateReg, "Wintop", IIf(Me.Top < 0, 0, Me.Top))
    Call SaveSetting("ZLSOFT", strPrivateReg, "WinWidth", Me.Width)
    Call SaveSetting("ZLSOFT", strPrivateReg, "WinHeight", Me.Height)
    Call SaveSetting("ZLSOFT", strPrivateReg, "MiniatureW", ucBgImages.Width)
    Call SaveSetting("ZLSOFT", strPrivateReg, "缩略图数量", ucBgImages.PageRecordCount)
End Sub

Private Function GetWindowCaption() As String
    Dim lngCurWinType As TImgProcessType
    
    lngCurWinType = mlngWinType
    
    If lngCurWinType = ptMark Then
        GetWindowCaption = "标记图"
    Else
        GetWindowCaption = IIf(lngCurWinType = ptPreview, "图像预览", "图像处理")
    End If
End Function

Private Sub RestorceWinLayout()
    Dim strCaption As String
    Dim strPrivateReg As String
     
    strCaption = GetWindowCaption()
    
    strPrivateReg = GetPrivateRegPath(strCaption)
    
    Me.Left = nvl(GetSetting("ZLSOFT", strPrivateReg, "WinLeft", Screen.Width / 4))
    Me.Top = nvl(GetSetting("ZLSOFT", strPrivateReg, "Wintop", Screen.Height / 4))
    Me.Width = nvl(GetSetting("ZLSOFT", strPrivateReg, "WinWidth", Screen.Width / 2))
    Me.Height = nvl(GetSetting("ZLSOFT", strPrivateReg, "WinHeight", Screen.Height / 2))

    ucBgImages.Width = nvl(GetSetting("ZLSOFT", strPrivateReg, "MiniatureW", 3000))
    
    ucBgImages.PageRecordCount = Val(GetSetting("ZLSOFT", strPrivateReg, "缩略图数量", 8))
End Sub


Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim objControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
    With Me.cbrMain.options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 16, 16
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    '图像操作工具栏定义
    Set cbrToolBar = Me.cbrMain.Add("图像操作栏", xtpBarTop)
'    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True '文本显示在图标下方
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        
        If mlngWinType = ptMark Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Process_Save, "保存"): cbrControl.ToolTipText = "保存"
        Else
            Set cbrControl = .Add(xtpControlButton, conMenu_Process_SaveToStudy, "存为检查图"): cbrControl.ToolTipText = "保存到检查图像"
            If mblnIsReportShow Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Process_SaveToReport, "存为报告图"): cbrControl.ToolTipText = "保存到报告图像"
            End If
        End If
        
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Window, "亮度"): cbrControl.ToolTipText = "调节亮度/对比度": cbrControl.Visible = mlngWinType <> ptMark
        cbrControl.Checked = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Zoom, "缩放"): cbrControl.ToolTipText = "缩放图像": cbrControl.Visible = mlngWinType <> ptMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RectZoom, "裁剪"): cbrControl.ToolTipText = "裁剪采集图像": cbrControl.iconid = 3201: cbrControl.Visible = mlngWinType <> ptMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RRotate, "顺时"): cbrControl.ToolTipText = "顺时针旋转": cbrControl.Visible = mlngWinType <> ptMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_LRotate, "逆时"): cbrControl.ToolTipText = "逆时针旋转": cbrControl.Visible = mlngWinType <> ptMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Sharpness, "锐化"): cbrControl.ToolTipText = "锐化": cbrControl.Visible = mlngWinType <> ptMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Filter, "平滑"): cbrControl.ToolTipText = "平滑": cbrControl.Visible = mlngWinType <> ptMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Line, "直线"): cbrControl.ToolTipText = "直线标注": cbrControl.Visible = mlngWinType <> ptMark
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Arrow, "箭头"): cbrControl.ToolTipText = "箭头标注": cbrControl.Visible = mlngWinType <> ptMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Ellipse, "圆形"): cbrControl.ToolTipText = "圆形标注"
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_Process_NumTag, "数字"): cbrControl.ToolTipText = "数字标注"
        Call LoadComNumber(cbrControl)
        
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_Process_TextTag, "文本"): cbrControl.ToolTipText = "常用文本标注"
        Call LoadComText(cbrControl)
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_MoveLabel, "移动"): cbrControl.ToolTipText = "选中标注时，鼠标左键拖拽移动标注，否则拖动图片，双击删除标注"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_LabelSetUp, "设置标注"): cbrControl.ToolTipText = "设置文字标注"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_DelAllLabels, "清除"): cbrControl.ToolTipText = "清除全部标注"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Restore, "恢复"): cbrControl.ToolTipText = "恢复图像到初始状态"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Exit, "退出"): cbrControl.ToolTipText = "退出"
    End With
    For Each cbrControl In cbrToolBar.Controls
         cbrControl.Style = xtpButtonIconAndCaption
         cbrControl.Category = "Main" '设置成主界面菜单
    Next
    cbrToolBar.Position = xtpBarTop
End Sub

Private Sub LoadComNumber(mnuParent As Object)
    Dim objControl As CommandBarControl
    Dim i As Long
    
    Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_Num * 100, "①"): objControl.ToolTipText = "自动递增数字编号": objControl.Category = 0: objControl.iconid = 0
    
    For i = 1 To 9
        Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_Num * 100 + i, i): objControl.ToolTipText = "数字编号" & i: objControl.Category = i: objControl.iconid = 0
    Next
    
End Sub

Private Sub LoadComText(mnuParent As Object)
    Dim objControl As CommandBarControl
    Dim arrTemp() As String
    Dim i As Long
    
    arrTemp = Split(mStrTemp, "|")
    
    For i = 0 To UBound(arrTemp)
        If Len(arrTemp(i)) > 0 Then
            Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_Word * 100 + i + 1, arrTemp(i)): objControl.ToolTipText = "文本标注": objControl.Category = i + 1: objControl.iconid = 0
        End If
    Next
End Sub


Private Sub lstMemoText_DblClick()
    cbxMemoText.Text = lstMemoText.list(lstMemoText.ListIndex)
    lstMemoText.Visible = False
    
    cbxMemoText.SelStart = 0
    cbxMemoText.SelLength = Len(cbxMemoText.Text)
    cbxMemoText.SetFocus
End Sub

Private Sub lstMemoText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        Call zlControl.CboSetText(cbxMemoText, lstMemoText.list(lstMemoText.ListIndex))
    End If
End Sub

Private Sub lstMemoText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlControl.CboSetText(cbxMemoText, lstMemoText.list(lstMemoText.ListIndex))
    End If
    
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then lstMemoText.Visible = False
End Sub

Private Sub picCboDropDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picCboDropDown.BorderStyle = 1
End Sub

Private Sub picCboDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picCboDropDown.BorderStyle = 0
End Sub
 

Private Sub picMemo_Resize()
    On Error Resume Next
    
    '摆放备注文字
    Me.lblMemoText.Left = 100
    Me.lblMemoText.Top = 200

    Me.cbxMemoText.Left = Me.lblMemoText.Left + Me.lblMemoText.Width
    Me.cbxMemoText.Top = Me.lblMemoText.Top - 100
    Me.cbxMemoText.Width = Me.ScaleWidth - Me.cbxMemoText.Left - 250 - cmdInsert.Width - cmdFont.Width - cmdAdd.Width
    
    Me.picCboDropDown.Left = Me.cbxMemoText.Left + Me.cbxMemoText.Width - 270
    Me.picCboDropDown.Top = Me.cbxMemoText.Top + 30
    
    Me.cmdAdd.Left = Me.cbxMemoText.Left + Me.cbxMemoText.Width
    Me.cmdAdd.Top = Me.cbxMemoText.Top
    
    Me.cmdInsert.Left = Me.cmdAdd.Left + Me.cmdAdd.Width
    Me.cmdInsert.Top = Me.cbxMemoText.Top

    Me.cmdFont.Left = Me.cmdInsert.Left + Me.cmdInsert.Width
    Me.cmdFont.Top = Me.cmdInsert.Top
    
    err.Clear
End Sub

Private Sub Timer1_Timer()
    Dim ptWin As POINTAPI
On Error GoTo errHandle
    If mlngState = 3 Then
        Timer1.Enabled = False
        Timer2.Enabled = False
        
        Exit Sub
    End If
    
    If mlngAdviceId = 0 Then
        Timer1.Enabled = False
        Exit Sub
    End If
    
    GetCursorPos ptWin

    If mlngState = 1 Then
        '鼠标移进窗体
        If ptWin.X >= Me.Left / 15 And ptWin.X <= (Me.Left + Me.Width) / 15 And ptWin.Y >= Me.Top / 15 And ptWin.Y <= (Me.Top + Me.Height) / 15 Then
 
            mlngState = 2
            
            If DViewer.Images.Count > 0 Then
                Call ClearHint(DViewer.Images(1))
            End If
            
            Call RefrshObjVisible
            Call refreshFace
            

            
            If ucBgImages.ImgCount <= 0 And ThumbnailImgCount > 0 Then
            '如果是预览，则在鼠标第一次移动到窗体时加载图像
                Call LoadImgs(maryImgInfos)
            End If
              
            Timer2.Enabled = False
        End If
    ElseIf mlngState = 2 Then
        '鼠标移出窗体

        If ptWin.X < Me.Left / 15 Or ptWin.X > (Me.Left + Me.Width) / 15 Or ptWin.Y < Me.Top / 15 Or ptWin.Y > (Me.Top + Me.Height) / 15 Then
            mlngState = 1
            
            If DViewer.Images.Count > 0 Then
                Call DrawHintTag(DViewer.Images(1))
            End If
            
            Call RefrshObjVisible
            Call refreshFace
   
            If mlngPreViewTime > 0 And mlngWinType = ptPreview Then
                Timer2.Enabled = True
            End If
            
            
        End If
    Else
        Timer1.Enabled = False
        Timer2.Enabled = False
    End If
Exit Sub
errHandle:
    Debug.Print "Timer1 Bug:" & err.Description
End Sub



Private Sub Timer2_Timer()
    If mlngWinType = ptPreview And mlngPreViewTime > 0 Then
        Call UnloadMe
    End If
End Sub

Private Sub txtInputText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 27 Then  '''ESC和回车键退出输入
        txtInputText.Visible = False
        If Trim(txtInputText.Text) = "" Then
            '删除文字标注
            DViewer.Images(1).Labels.Remove DViewer.Images(1).Labels.Count
            txtInputText = "1 "
        Else
            mdcmSelectLabel.Text = txtInputText.Text
            DViewer.Refresh
        End If
    End If
End Sub

Private Sub ShowPopupImage()
'------------------------------------------------
'功能：创建鼠标右键弹出菜单
'intType:0--报告图，1--缩略图，2--缓存图
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Page, "分页设置")
            
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub

Private Sub subaCorrectCursor(v As DicomViewer, im As DicomImage, xx As Long, Yy As Long)
'------------------------------------------------
'功能：鼠标移动如果超出图像范围则修正其鼠标位置
'参数：v--图像所在的viewer；im--鼠标所在的图像；xx--鼠标所在的x方向位置，如果鼠标超出图像则将此值修改到图像之内；
'      yy--鼠标所在的y方向位置，如果鼠标超出图像则将此值修改到图像之内；
'返回：无
'------------------------------------------------
    Dim X As Integer, Y As Integer, w As Long, h As Long
    Dim i As DicomImage
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    w = v.Width / v.MultiColumns / Screen.TwipsPerPixelX - v.CellSpacing * 2
    h = v.Height / v.MultiRows / Screen.TwipsPerPixelY - v.CellSpacing * 2
    X = im.OriginX + v.CellSpacing
    Y = im.OriginY + v.CellSpacing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If xx < X Then xx = X
    If xx > X + w Then xx = X + w
    If Yy < Y Then Yy = Y
    If Yy > Y + h Then Yy = Y + h
End Sub

Public Sub subMoveLable(la As DicomLabel, X As Long, Y As Long)
'------------------------------------------------
'功能：移动一个标注
'参数：la--被移动的标注；x--x方向移动的图像像素距离；y--y方向移动的图像像素距离
'返回：无
'------------------------------------------------
    
    la.Left = la.Left + X
    la.Top = la.Top + Y
    
    '如果是数字编号，需要同时移动三个标注
    If la.tag <> "" And Not la.TagObject Is Nothing Then
        la.TagObject.Left = la.TagObject.Left + X
        la.TagObject.Top = la.TagObject.Top + Y
        la.TagObject.TagObject.Left = la.TagObject.TagObject.Left + X
        la.TagObject.TagObject.Top = la.TagObject.TagObject.Top + Y
    End If
       
End Sub

Private Sub subSetTextLabel()
'------------------------------------------------
'功能：设置文字标注，并保存
'参数：
'返回：无
'------------------------------------------------
    Dim strTemp As String
    Dim i As Integer

    On Error GoTo err
    
'    strTemp = InputBox("请输入新的文字标注配置，格式为“简码1=说明1|简码2=说明2|...”。", "文字标注设置", Replace(mstrTemp, "[+]", "|"))
    
    strTemp = frmInputBoxV2.ZlShowMe(Me, mStrTemp)
    
    
    If strTemp = "" Then Exit Sub
    
    If InStr(strTemp, "=") = 0 Then
        MsgBoxD Me, "输入的格式不正确，应该按照“简码=说明”方式输入，请检查后重新设置。", vbOKOnly, "提示"
        Exit Sub
    End If
     
    '输入成功，使用这个新的文字标注，同时保存到注册表中
    mStrTemp = strTemp

    cbrMain.FindControl(, conMenu_Process_TextTag).CommandBar.Controls.DeleteAll
    
    LoadComText cbrMain.FindControl(, conMenu_Process_TextTag)
    Call SaveSetting("ZLSOFT", "公共模块\zl9PACSWork\frmReportImageEdit", "简明文字标注", Replace(mStrTemp, "|", "[+]"))
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subLoadTextLabel()
'------------------------------------------------
'功能：读取文字标注
'参数：
'返回：无
'------------------------------------------------
    Dim strTemp As String
    Dim strtext() As String
    Dim i  As Integer
    
    On Error GoTo err
    
    mStrTemp = GetSetting("ZLSOFT", "公共模块\zl9PACSWork\frmReportImageEdit", "简明文字标注", G_STR_TAG)
    
    mStrTemp = Replace(mStrTemp, "[+]", "|")
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub



Public Sub UnloadMe()
    Unload Me
End Sub

 

Private Sub ucBgImages_OnClick(ByVal lngSelIndex As Long)
On Error GoTo err
     
    Call EnterProcessState
    
    If mlngWinType = ptMark Then Exit Sub
    
 
    If DViewer.Images.Count > 0 Then
        If mblnIsChanged Then
            If MsgBoxD(Me, "图像操作尚未保存，是否继续？", vbYesNo, "提示") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    DViewer.Images.Clear
    
    Set mOldImage = ucBgImages.GetImage(lngSelIndex)
    
    If mOldImage Is Nothing Then Exit Sub
    DViewer.Images.Add ucBgImages.GetImage(lngSelIndex)
     
    mblnIsChanged = False
    mblnCase = False
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub


Public Sub AutoUnload()
    If mblnIsUnloud Then
        Timer2.Enabled = True
    End If
End Sub

Private Sub DrawHintTag(dcmImg As DicomImage)
    Dim lRpt As DicomLabel
    Dim i As Integer
    
    If mlngAdviceId = 0 Then Exit Sub
     
    Set lRpt = New DicomLabel
            
    With lRpt
        .LabelType = doLabelText
        .Width = 800
        .Height = 60
        .ImageTied = False
        .Transparent = True
        .ScaleWithCell = True
        .ScaleFontSize = 40
        .Font.Name = "宋体"
        .Font.Size = 22
        .Font.Bold = True
        .ForeColour = &HCBBECB
        .Left = 120
        .Top = 20
        .Text = "...更多操作请点击..."
        .Shadow = doShadowBottomRight
        .Alignment = doAlignCentre
        .Visible = True
        .tag = "HINT"
    End With
    
    dcmImg.Labels.Add lRpt
    
    dcmImg.Refresh False
End Sub

Private Sub ClearHint(dcmImage As DicomImage)
    Dim i As Long
    
    For i = 1 To dcmImage.Labels.Count
        If dcmImage.Labels(i).tag = "HINT" Then
            dcmImage.Labels.Remove i
            Exit For
        End If
    Next
    
    dcmImage.Refresh False
End Sub
 

Private Sub ucSplitter_OnMoveEnd()
    If lstMemoText.Visible Then lstMemoText.ZOrder
End Sub

