VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImageProcess 
   Caption         =   "图像处理"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11550
   Icon            =   "frmImageProcess.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11550
   StartUpPosition =   3  '窗口缺省
   Begin zl9PacsControl.ucSplitter ucSplitter 
      Height          =   6735
      Left            =   3015
      TabIndex        =   11
      Top             =   240
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   11880
      Control1Name    =   "picImage"
      Control2Name    =   "DViewer"
   End
   Begin VB.ListBox lstMemoText 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3200
      Left            =   9480
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox picImage 
      Height          =   6735
      Left            =   120
      ScaleHeight     =   6675
      ScaleWidth      =   2835
      TabIndex        =   8
      Top             =   240
      Width           =   2895
      Begin zl9PacsControl.ucImageThumbnail ucMiniature 
         Height          =   6255
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   11033
         ShowCheckbox    =   -1  'True
      End
      Begin zl9PacsControl.ucSplitPageNew ucPage 
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   6360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         PageCount       =   0
         PageRecord      =   6
         AutoRedrawStyle =   0   'False
      End
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
         Picture         =   "frmImageProcess.frx":6852
         ScaleHeight     =   375
         ScaleWidth      =   255
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   120
         Width           =   5655
      End
      Begin VB.CommandButton cmdFont 
         Height          =   375
         Left            =   9000
         Picture         =   "frmImageProcess.frx":6BAE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "设置当前备注字体。"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdInsert 
         Height          =   375
         Left            =   8640
         Picture         =   "frmImageProcess.frx":6EF0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "将当前备注设置为常用备注"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   375
         Left            =   8280
         Picture         =   "frmImageProcess.frx":765A
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
      Left            =   4680
      Top             =   7560
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
      Height          =   6735
      Left            =   3150
      TabIndex        =   7
      Top             =   240
      Width           =   7815
      _Version        =   262147
      _ExtentX        =   13785
      _ExtentY        =   11880
      _StockProps     =   35
      BackColor       =   -2147483638
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
Attribute VB_Name = "frmImageProcess"
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

'Private mlngModule As Long
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

Private mblnOK As Boolean
Private mOldImage As DicomImage
Private mlngImgIndex As Long            '父窗体选中缩略图的索引
Private mblnIsMark As Boolean           '是标记图
Private mintTextIndex As Integer        '文字标注按钮的索引
Private mstrText As String              '文字标注内容
Private mstrCustom  As String           '自定义标注内容
Private mintNumberIndex As Integer      '数字编号按钮的索引
Private mintAutoNumber As Integer       '自动递增编号的最大号码
Private mstrTemp As String
Private mstrUser As String
Private mblnPreView As Boolean          '是否预览
Private mlngWinType As Long             '打开窗口时窗口类型
Private mlngPreViewTime As Long         '移动预览延时关闭时间
Private mlngState As Long               '预览图像窗口状态，1-预览；2-处理；3-单击后
Private mobjDownLoadImages As New clsImageDownload
Private mobjService As New clsServiceHelper
Private mblnMoved As Boolean
Private mstrQueryValue As String
Private mblnIsUnloud As Boolean         '当前鼠标位置是否自动关闭
Private mblnDrag As Boolean
Private mintDisState As Integer
Private mblnIsChanged As Boolean
Private mblnCase As Boolean
Private mblnDoShiled As Boolean

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
Public Event AfterSaveStady()
Public Event OnSaveImage(ByVal dcmImage As DicomImage, ByVal emImageType As TImageType)


Public Sub zlShowMe(ByVal strQueryValue As String, dcmImage As DicomImage, _
    lngImgIndex As Long, objParent As Object, blnRefresh As Boolean, blnMoved As Boolean, Optional lngLeval As Long, Optional lngType As Long = 0, Optional lngPreviewTime As Long = 0, Optional blnDoShiled As Boolean)
'lngType:窗口类型，0-图像处理窗口；1-图像预览窗口；2-标记图处理窗口
    
    On Error GoTo err
    
    Dim i As Integer
    Dim arrImages() As String
    
    mblnPreView = lngType = 1
    mblnIsMark = lngType = 2
    mlngWinType = lngType
    mblnMoved = blnMoved
    mstrQueryValue = strQueryValue
    mlngPreViewTime = lngPreviewTime
    mblnDoShiled = blnDoShiled
    mstrUser = GetUserInfo
    
    If IsChanged Then
        If MsgBox("有尚未保存的图像处理，该操作将清空这些处理，是否继续？", vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        Else
            Call InitChangedState
        End If
    End If
    
    mblnDrag = False
    mblnIsChanged = False
    mblnCase = False
    
    Timer1.Enabled = False
    Timer2.Enabled = False
    
    If lngType = 1 Then mlngState = 1

    If Not mblnIsMark Then
        ucMiniature.ShowCheckBox = False
        
        mlngImgIndex = mobjDownLoadImages.GetImageIdex(lngLeval, strQueryValue, False, blnMoved, dcmImage.InstanceUID)
        mobjDownLoadImages.ImgLoadType = IIf(mobjService.GetServiceStatus = SERVICE_RUNNING, FileLoadType.Service, FileLoadType.Normal)
        mobjDownLoadImages.QueryLevel = lngLeval
        
        If Not blnRefresh Then
            ucPage.RecordCount = mobjDownLoadImages.GetRecordCount(lngLeval, strQueryValue, False, blnMoved)
            
            Call mobjDownLoadImages.DownloadImages(arrImages, strQueryValue, IIf(ucPage.PageNumber = 0, 1, (ucPage.PageNumber - 1) * ucPage.PageRecord + 1), IIf(ucPage.PageNumber = 0, 1, ucPage.PageNumber) * ucPage.PageRecord, False, blnMoved)
            Call ucMiniature.SplitPage(ucPage)
            Call ucMiniature.RefreshImage(arrImages())
        End If
        
        If mlngImgIndex > 0 And mlngImgIndex <= ucPage.RecordCount Then
            ucPage.MoveItem (mlngImgIndex)
        ElseIf mlngImgIndex = 0 And ucPage.RecordCount > 0 Then
            ucPage.MoveItem 1
        End If
        
    End If
    
    If mblnIsMark Then
        Set mOldImage = dcmImage
            
        Me.DViewer.Images.Clear
        Me.DViewer.Images.Add dcmImage
    End If
    
    If DViewer.Images.Count > 0 Then
        ClearLable Me.DViewer.Images(1)
    End If
    
    '重建标注之间的关联
    If DViewer.Images.Count > 0 Then
        Call subLabelCopyRebuild(dcmImage, Me.DViewer.Images(1))
    End If
    
    If Not blnRefresh Then
        Me.Show 0, objParent
        SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3  '将窗口置顶
    End If
    
    Call RefrshObjVisible
    
    
    If lngType = 1 Then
        If DViewer.Images.Count > 0 Then
            Call DrawHintTag(DViewer.Images(1))
        End If
            
        Timer1.Enabled = True

        If lngPreviewTime > 0 Then
            Timer2.Interval = lngPreviewTime * 1000
            Timer2.Enabled = True
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub RefrshObjVisible()
    If Not mblnIsMark Then
        Me.ucMiniature.Visible = Not mblnPreView
        Me.cbxMemoText.Visible = Not mblnPreView
        Me.cmdAdd.Visible = Not mblnPreView
        Me.cmdFont.Visible = Not mblnPreView
        Me.cmdInsert.Visible = Not mblnPreView
        Me.picImage.Visible = Not mblnPreView
        Me.ucSplitter.Visible = Not mblnPreView
        Me.picMemo.Visible = Not mblnPreView

        If Me.lstMemoText.Visible Then
            Me.lstMemoText.Visible = Not mblnPreView
        End If
        Me.ucPage.Visible = Not mblnPreView

        If Me.txtInputText.Visible Then
            Me.txtInputText.Visible = Not mblnPreView
        End If
        Me.picCboDropDown.Visible = Not mblnPreView
        Me.lblMemoText.Visible = Not mblnPreView

        Me.cbrMain.FindControl(, conMenu_Process_Window).Parent.Visible = Not mblnPreView
        
        Me.Caption = IIf(mblnPreView, "图像预览", "图像处理")
    Else

        Me.ucMiniature.Visible = Not mblnIsMark

        Me.lblMemoText.Visible = Not mblnIsMark
        Me.cbxMemoText.Visible = Not mblnIsMark
        Me.picCboDropDown.Visible = Not mblnIsMark
        Me.cmdInsert.Visible = Not mblnIsMark
        Me.cmdFont.Visible = Not mblnIsMark
        Me.cmdAdd.Visible = Not mblnIsMark
        Me.ucPage.Visible = Not mblnIsMark
        Me.picImage.Visible = Not mblnIsMark
        Me.ucSplitter.Visible = Not mblnIsMark
        Me.picMemo.Visible = Not mblnIsMark
    End If
End Sub

Private Sub ClearLable(dcmImage As DicomImage)
    Dim i As Long
     '去除边框
    For i = 1 To dcmImage.Labels.Count
        If dcmImage.Labels(i).Tag = "SELECT" Or dcmImage.Labels(i).Tag = "BORDER" Or dcmImage.Labels(i).Tag = "HINT" Then
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
        If InStr(Trim(NVL(mrsTmp!简码)), UCase(cbxMemoText.Text)) > 0 Or InStr(Trim(NVL(mrsTmp!名称)), UCase(cbxMemoText.Text)) > 0 Then
            getListIndex = i
            
            Exit For
        End If

        mrsTmp.MoveNext
    Next
End Function

Private Sub cbxMemoText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        If lstMemoText.ListCount > 0 Then lstMemoText.ListIndex = getListIndex
    End If
End Sub
'
'Private Sub cmdNum_Click(Index As Integer)
'    mintNumberIndex = Index
'    subSetMouseState msNumber
'
'    Call setCmdLabelColor
'
'    cmdNum(Index).BackColor = &HC0C000
'End Sub

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
    
    If mlngWinType = 1 Then
        mlngState = 3
        Timer1.Enabled = False
        Timer2.Enabled = False
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
End Sub

Private Function GetNewImage(emImageType As TImageType) As DicomImage
    Dim dcmGlobal As New DicomGlobal
    Dim dcmImage As DicomImage
    Dim Img As New DicomImage
    Dim iPlane As Integer
    Dim dcmTag As clsImageTagInf
    
    dcmGlobal.RegString("UIDRoot") = "1"
    
  
    If Me.DViewer.Images.Count = 1 Then
        Set dcmImage = Me.DViewer.Images(1)
        If emImageType <> mtTagImage Then
'            Set GetNewImage = dcmImage.PrinterImage(8, iPlane, True, 1, 0, dcmImage.SizeX, 0, dcmImage.SizeY)
            
            '转换一次图片格式，保存标注
            dcmImage.FileExport App.Path & "\PacsControlCacheImg.jpg", "JPG"
            Img.FileImport App.Path & "\PacsControlCacheImg.jpg", "JPG"
            Kill App.Path & "\PacsControlCacheImg.jpg"
            
            Set GetNewImage = Img
            GetNewImage.InstanceUID = dcmGlobal.NewUID
            GetNewImage.SeriesUID = dcmImage.SeriesUID
            GetNewImage.StudyUID = dcmImage.StudyUID
            
            If emImageType = mtReportImage Then
                GetNewImage.BorderWidth = 1
                GetNewImage.BorderColour = vbWhite
            End If
            '设置图像标记
            Set dcmTag = New clsImageTagInf
            dcmTag.Tag = imgTag
            
            Set GetNewImage.Tag = dcmTag
        Else
            dcmImage.InstanceUID = dcmGlobal.NewUID
            
            Set GetNewImage = dcmImage
        End If
    Else
        Set GetNewImage = Nothing
    End If
End Function

'Private Sub cmdReport_Click()
''功能：添加操作但不关闭窗口
''参数：
''返回：无
''------------------------------------------------
'    On Error GoTo err
'
'
'err:
'    If ErrCenter() = 1 Then Resume Next
'End Sub


Private Sub SaveImage(emImageType As TImageType)
    Dim dcmImage As DicomImage
    
    Set dcmImage = GetNewImage(emImageType)
    If dcmImage Is Nothing Then Exit Sub
    
    RaiseEvent OnSaveImage(dcmImage, emImageType)
    
    If Not ucMiniature.SelectImage Is Nothing Then
        ucMiniature.SelectImage.Tag.IsChanged = False
    End If
    mblnIsChanged = False
    If emImageType = mtStadyImage Then
        Call AfterSaveStudy(dcmImage)
    End If
    
    If emImageType = mtTagImage Then
        Unload Me
    End If
End Sub


Private Sub DViewer_Click()
    On Error GoTo err
    
    If mlngWinType = 1 Then
        mlngState = 3
        Timer1.Enabled = False
        Timer2.Enabled = False
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

'Private Sub cmdTextLabel_Click(Index As Integer)
'
'    If Index = 11 Then '自定义，先判断是否有输入文字
'        If Trim(txtUserLabelText.Text) = "" Then
'            MsgBox "请输入自定义标注。", vbOKOnly, CON_STR_HINT_TITLE
'            txtUserLabelText.SetFocus
'            Exit Sub
'        End If
'    End If
'    mintTextIndex = Index
'    subSetMouseState msFixText
'
'    Call setCmdLabelColor
'
'    cmdTextLabel(Index).BackColor = &HC0C000
'End Sub

'Private Sub setCmdLabelColor()
'    Dim i As Integer
'
'    For i = 0 To cmdTextLabel.Count - 1
'        cmdTextLabel(i).BackColor = &H8000000F
'    Next i
'
'    For i = 0 To cmdNum.Count - 1
'        cmdNum(i).BackColor = &H8000000F
'    Next i
'End Sub

Private Sub DViewer_DblClick()
    Dim ls As DicomLabels
    Dim l As DicomLabel
    
    On Error GoTo err
    
    If mintMouseState = msMove Then
        Set ls = DViewer.LabelHits(mlngBaseXX, mlngBaseYY, False, False, True)
        If ls.Count > 0 Then
            If MsgBox("是否删除这个标注？", vbOKCancel, CON_STR_HINT_TITLE) = vbOK Then
                Set l = ls(1)
                If l.Tag <> "" Then
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

Private Sub DViewer_MouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)
    Call ucMiniature.MouseWheel(Delta)
End Sub

Private Sub Form_Resize()
'    Call RefreshFace
End Sub

Private Sub lstMemoText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call zlControl.CboSetText(cbxMemoText, lstMemoText.List(lstMemoText.ListIndex))
End Sub

Private Sub picCboDropDown_Click()
    lstMemoText.ZOrder
    lstMemoText.Visible = Not lstMemoText.Visible
    If lstMemoText.ListCount > 0 Then lstMemoText.ListIndex = getListIndex

    If lstMemoText.Visible Then lstMemoText.SetFocus
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

'Private Sub cmdCur_Click()
''上一幅图像
'On Error GoTo errH
'
'    Call ChangeImage(1)
'
'    Exit Sub
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

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
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

'Private Sub cmdNext_Click()
''下一幅图像
'On Error GoTo errH
'
'    Call ChangeImage(2)
'
'    Exit Sub
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

'Private Sub cmdAdd_Click()
''------------------------------------------------
''功能：添加操作但不关闭窗口
''参数：
''返回：无
''------------------------------------------------
'    On Error GoTo err
'
'    Dim dcmGlobal As New DicomGlobal
'
'    dcmGlobal.RegString("UIDRoot") = "1"
'    mblnOK = True
'    '拼接方法
'    Call subAddMemoText
'
'    If mblnOK Then
'        If Me.DViewer.Images.Count = 1 Then
'            Set mImage = Me.DViewer.Images(1)
'            mImage.InstanceUID = dcmGlobal.NewUID   '图像处理创建图像后，就设置新的InstanceUID
'        Else
'            Set mImage = Nothing
'        End If
'    Else
'        Set mImage = Nothing
'    End If
'
'    If mblnIsMark = True Then   '标记图处理，添加标记图后直接退出
'        Call mfrmParent.DcmAddMarkImage(mImage)
'        Unload Me
'        Exit Sub
'    Else    '采集的图像处理
'        '对拼接后的图像的边框进行处理
'         If Me.DViewer.Images.Count > 0 Then
'             With Me.DViewer.Images(1)
'                .BorderWidth = 3
'                .BorderStyle = 2
'                .BorderColour = vbRed
'            End With
'        End If
'
'        Call mfrmParent.DcmAddImage(mImage, mSelViewerIndex)
'    End If
'
'    Me.DViewer.Refresh
'
'    '清空ComboBox文本
'    cbxMemoText.Text = ""
'
'    '关闭下拉框
'    lstMemoText.Visible = False
'    Exit Sub
'err:
'    If ErrCenter() = 1 Then Resume Next
'End Sub

'Private Sub cmdExit_Click()
''清空Viewer控件，并卸载窗口
'   ' Me.DViewer.Images.Clear
'    Unload Me
'End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim lngCount As Long
    
    On Error GoTo errHandle
    
    If mlngWinType = 1 Then
        mlngState = 3
        Timer1.Enabled = False
        Timer2.Enabled = False
    End If
    
    mblnDrag = False
    Select Case control.ID
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
            mintNumberIndex = Val(control.Category)
            subSetMouseState msNumber
        
        Case conMenu_Process_NumTag
            mintNumberIndex = 0
            subSetMouseState msNumber
            
        Case conMenu_Process_Word * 100 To conMenu_Process_Word * 100 + 99
            mstrText = control.Caption
            subSetMouseState msFixText
        
        Case conMenu_Process_Save           '标记图保存
            If mblnIsMark Then
                Call SaveImage(mtTagImage)
            End If
            
        Case conMenu_Process_SaveToStady         '保存到检查
            Call SaveImage(mtStadyImage)
            
            
        Case conMenu_Process_SaveToReport           '保存到报告
            Call SaveImage(mtReportImage)
            
        Case conMenu_Process_Exit             '退出
            Unload Me
    End Select
    
'    If Control.ID <> conMenu_Process_LabelSetUp Then
'        Call setCmdLabelColor
'    End If
    
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
    Dim Img As DicomImage
    Dim lblFrame As DicomLabel
    
    If Me.DViewer.Images.Count <> 1 Then Exit Sub
    If Me.DViewer.Images(1).Labels.Count < 1 Then Exit Sub
    
    Set Img = Me.DViewer.Images(1)
    Set lblFrame = Me.DViewer.Images(1).Labels(Me.DViewer.Images(1).Labels.Count)
    
    If Abs(lblFrame.Width) = 0 Or Abs(lblFrame.Height) = 0 Then
        MsgBox "请选择图像区域后再保存", vbExclamation, CON_STR_HINT_TITLE
        Exit Sub
    End If
    
    '图象最大宽高=300
    iMax = 300
    
    '根据label来提取被框选中的图像
    '图象位数,黑白图像为1，彩色图像为3
    iPlane = 1
    If Not IsNull(Img.Attributes(&H28, &H4).value) And Img.Attributes(&H28, &H4).Exists Then
        If Img.Attributes(&H28, &H4).value = "RGB" Or Img.Attributes(&H28, &H4).value = "YBR_FULL_422" Then
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
    
    Img.Labels(Img.Labels.Count).Visible = False
    If (Img.RotateState = doRotateLeft And Img.FlipState = doFlipNormal) _
        Or (Img.RotateState = doRotateRight And Img.FlipState = doFlipBoth) _
        Or (Img.RotateState = doRotate180 And Img.FlipState = doFlipVertical) _
        Or (Img.RotateState = doRotateNormal And Img.FlipState = doFlipHorizontal) Then
        'X方向对调
        Set imgResult = Img.PrinterImage(8, iPlane, True, dblZoom, Img.SizeX - iRight, Img.SizeX - iLeft, iTop, iBottom)
    ElseIf (Img.RotateState = doRotateLeft And Img.FlipState = doFlipBoth) _
        Or (Img.RotateState = doRotateRight And Img.FlipState = doFlipNormal) _
        Or (Img.RotateState = doRotate180 And Img.FlipState = doFlipHorizontal) _
        Or (Img.RotateState = doRotateNormal And Img.FlipState = doFlipVertical) Then
        'Y方向对调
        Set imgResult = Img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, Img.SizeY - iBottom, Img.SizeY - iTop)
    ElseIf (Img.RotateState = doRotateRight And Img.FlipState = doFlipHorizontal) _
        Or (Img.RotateState = doRotateLeft And Img.FlipState = doFlipVertical) _
        Or (Img.RotateState = doRotate180 And Img.FlipState = doFlipNormal) _
        Or (Img.RotateState = doRotateNormal And Img.FlipState = doFlipBoth) Then
        'X，Y方向对调
        Set imgResult = Img.PrinterImage(8, iPlane, True, dblZoom, Img.SizeX - iRight, Img.SizeX - iLeft, Img.SizeY - iBottom, Img.SizeY - iTop)
    Else
        Set imgResult = Img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, iTop, iBottom)
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
    Call RefreshFace
     
End Sub

Public Sub RefreshFace()
    '设置显示的客户区域
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    On Error Resume Next
    
    cbrMain.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    
    If Not (mblnIsMark Or mblnPreView) Then
        
        Me.ucSplitter.Top = lngTop
        Me.ucSplitter.Height = Abs(lngBottom - lngTop - 600)
        
        Me.picImage.Left = lngLeft
        Me.picImage.Top = lngTop
        Me.picImage.Height = Abs(lngBottom - lngTop - 600)

        '摆放DViewer
        Me.DViewer.Left = lngLeft + Me.picImage.Width + Me.ucSplitter.Width
        Me.DViewer.Top = lngTop
        Me.DViewer.Width = Abs(lngRight - lngLeft - Me.picImage.Width - ucSplitter.Width)
        Me.DViewer.Height = Abs(lngBottom - lngTop - 600)
        
        ucSplitter.RePaint
        If lstMemoText.Visible Then lstMemoText.ZOrder
    Else
        Me.picImage.Left = lngLeft - Me.picImage.Width - Me.ucSplitter.Width
        Me.DViewer.Left = lngLeft
        Me.DViewer.Top = lngTop
        Me.DViewer.Width = lngRight
        Me.DViewer.Height = Abs(lngBottom - lngTop)
        
        
    End If
    
    Me.picMemo.Left = lngLeft
    Me.picMemo.Top = Me.DViewer.Top + Me.DViewer.Height
    Me.picMemo.Height = 600
    Me.picMemo.Width = lngRight
    
    Me.lstMemoText.Left = Me.cbxMemoText.Left
    Me.lstMemoText.Top = Me.picMemo.Top + Me.cbxMemoText.Top - Me.lstMemoText.Height
    Me.lstMemoText.Width = Me.cbxMemoText.Width - 10
    
'    Me.cmdExit.Left = Me.ScaleWidth - Me.cmdExit.Width * 1.8
'    Me.cmdExit.Top = Me.cbxMemoText.Top + Me.cbxMemoText.Height + 300
'
'    Me.cmdStady.Left = Me.cmdExit.Left - Me.cmdStady.Width - 150
'    Me.cmdStady.Top = Me.cmdExit.Top
'
'    Me.cmdReport.Left = Me.cmdStady.Left - Me.cmdReport.Width - 150
'    Me.cmdReport.Top = Me.cmdStady.Top
End Sub


'Private Sub InitFaceScheme()
'    '初始界面布局
'    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, pane4 As Pane
'
'    With dkpMain
'        .CloseAll
'        .Options.HideClient = True
'        .Options.UseSplitterTracker = fale '实时拖动
'        .Options.ThemedFloatingFrames = True
'        .Options.AlphaDockingContext = True
'
'    End With
'
'    picImage.Visible = False
'    If Not (mblnIsMark Or mblnPreView) Then
'        picImage.Visible = True
'
'        Set Pane1 = dkpMain.CreatePane(1, mdblMiniatureW, picBox.Height, DockLeftOf, Nothing)
'
'        Pane1.Title = "缩略图"
'        Pane1.Handle = picBox.hwnd
'    End If
'
'    Set Pane2 = dkpMain.CreatePane(2, mlngDViewerW, picBox.Height, DockRightOf, Pane1)
'    Pane2.Title = "预览图"
'    Pane2.Handle = DViewer.hwnd
'End Sub


Private Sub cmdInsert_Click()
    Dim strSQL As String, i As Integer
    Dim strUser As String
    
    If Trim(cbxMemoText.Text) = "" Then
        MsgBox "请输入备注内容。", vbInformation, CON_STR_HINT_TITLE
        If cbxMemoText.Enabled Then cbxMemoText.SetFocus
        Exit Sub
    End If
    If cbxMemoText.ListIndex <> -1 Then
        MsgBox "该备注内容已经在常用备注中。", vbInformation, CON_STR_HINT_TITLE
        If cbxMemoText.Enabled Then cbxMemoText.SetFocus
        Exit Sub
    Else
        For i = 0 To cbxMemoText.ListCount - 1
            If UCase(Trim(cbxMemoText.List(i))) = UCase(Trim(cbxMemoText.Text)) Then
                MsgBox "该备注容已经在常用备注中。", vbInformation, CON_STR_HINT_TITLE
                If cbxMemoText.Enabled Then cbxMemoText.SetFocus
                Exit Sub
            End If
        Next
    End If
        
    On Error GoTo errH
    
    strSQL = zlCommFun.zlGetSymbol(cbxMemoText.Text)
    strSQL = "zl_影像图像备注_Insert('" & Replace(cbxMemoText.Text, "'", "''") & "','" & strSQL & "','" & mstrUser & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    AddComboItem cbxMemoText.hWnd, CB_ADDSTRING, 0, cbxMemoText.Text
    lstMemoText.AddItem cbxMemoText.Text
    MsgBox "已设置为常用备注。", vbInformation, CON_STR_HINT_TITLE
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
    
    Dim Img As DicomImage
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
    Set Img = Me.DViewer.Images(1)
    
    iLeft = 0
    iTop = 0
    iWidth = Img.SizeX
    iHeight = Img.SizeY + lngFontHeight

    '使用PrinterImage方法，可以将图像上的标签及标注同时进行绘制
    Set imgResult = Img.PrinterImage(8, iPlane, True, 1, 0, iWidth, 0, iHeight - lngFontHeight)
'

    '添加标注
    Dim dlMemoText As New DicomLabel
    
    dlMemoText.LabelType = doLabelText
    dlMemoText.ImageTied = True
    dlMemoText.Transparent = False
    dlMemoText.AutoSize = False
    dlMemoText.Left = 0
    dlMemoText.Top = Img.SizeY
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
    
    If mblnPreView Then Exit Sub
    
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
                If mMovingLabel.Tag <> "" Then
                    If mMovingLabel.Tag = m_LabelTag_Back Then
                        Set mMovingLabel = mMovingLabel.TagObject
                    ElseIf mMovingLabel.Tag = m_LabelTag_Circle Then
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
                        If mblnIsMark = False Then
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
    
    If mblnPreView Then Exit Sub
    
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
    If mblnPreView Then Exit Sub

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
            mdcmSelectLabel.Tag = m_LabelTag_Back
            
            '添加顺序编号圆形的两个附加标注，圆形框和数字
            DViewer.Images(1).Labels.Add GetNewLabel(doLabelEllipse, mdcmSelectLabel.Left, mdcmSelectLabel.Top, 14, 14)
            Set mdcmSelectLabel = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count)
            mdcmSelectLabel.XOR = False
            mdcmSelectLabel.ForeColour = vbBlack
            mdcmSelectLabel.Transparent = True
            mdcmSelectLabel.Tag = m_LabelTag_Circle
            mdcmSelectLabel.TagObject = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count - 1)
            
            DViewer.Images(1).Labels.Add GetNewLabel(doLabelText, mdcmSelectLabel.Left + 1, mdcmSelectLabel.Top, 0, 0)
            Set mdcmSelectLabel = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count)
            mdcmSelectLabel.ForeColour = vbBlack
            mdcmSelectLabel.Transparent = True
            mdcmSelectLabel.XOR = False
            mdcmSelectLabel.Tag = m_LabelTag_Number
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


Private Sub subCenterZoom(Img As DicomImage, Viewer As DicomViewer, dblZoom As Double, corpSize As TPoint)
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
    Img.Zoom = dblZoom
    Img.StretchToFit = False

            
    Img.ScrollX = (Img.SizeX * Img.ActualZoom - ScaleX(Viewer.Width, vbTwips, vbPixels) / Viewer.MultiColumns) / 2 + corpSize.X
    Img.ScrollY = (Img.SizeY * Img.ActualZoom - ScaleY(Viewer.Height, vbTwips, vbPixels) / Viewer.MultiRows) / 2 + corpSize.Y
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
    mblnOK = False
    mintAutoNumber = 0
    
    '图像处理，鼠标左键默认是调窗；图像标注，鼠标左键默认是移动标注
'    If mblnIsMark = True Then
'        Call subSetMouseState(msMove)
'    Else
'        Call subSetMouseState(msWinLevel)
'    End If
'
    Call ReadEnjoin
    
    Call RefreshFace
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

'载入备注字体样式
Private Sub LoadMemoFontStyle()
    Dim strFontStyle As String
    Dim aryFontStyle() As String
    
    '“宋体,12,B,U,S,I”
    
    strFontStyle = zlDatabase.GetPara("图像备注字体", glngSys, glngMoudle, "")
    
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

    Call zlDatabase.SetPara("图像备注字体", strFontStyle, glngSys, glngMoudle)
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
        AddComboItem cbxMemoText.hWnd, CB_ADDSTRING, 0, mrsTmp!名称
        
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
    If IsChanged Then
        If MsgBox("有处理的图片尚未保存，是否继续退出？", vbYesNo, Me.Caption) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    
    Call SaveWinLayout
    
    Call SaveMemoFontStyle
    
    RaiseEvent OnUnload
End Sub

Private Sub SaveWinLayout()
'保存窗体位置及界面布局
'由于默认窗口大小原因未使用ZL9COMLIB中的方法
    Call SaveSetting("ZLSOFT", "私有模块\" & mstrUser & "\界面设置\" & App.EXEName & "\" & Me.Name, "WinLeft", IIf(Me.Left < 0, 0, Me.Left))
    Call SaveSetting("ZLSOFT", "私有模块\" & mstrUser & "\界面设置\" & App.EXEName & "\" & Me.Name, "Wintop", IIf(Me.Top < 0, 0, Me.Top))
    Call SaveSetting("ZLSOFT", "私有模块\" & mstrUser & "\界面设置\" & App.EXEName & "\" & Me.Name, "WinWidth", Me.Width)
    Call SaveSetting("ZLSOFT", "私有模块\" & mstrUser & "\界面设置\" & App.EXEName & "\" & Me.Name, "WinHeight", Me.Height)
    Call SaveSetting("ZLSOFT", "私有模块\" & mstrUser & "\界面设置\" & App.EXEName & "\" & Me.Name, "MiniatureW", picImage.Width)
    Call SaveSetting("ZLSOFT", "私有模块\" & mstrUser & "\界面设置\" & App.EXEName & "\" & Me.Name, "PageRecord", ucPage.PageRecord)
End Sub

Private Sub RestorceWinLayout()
    Me.Left = NVL(GetSetting("ZLSOFT", "私有模块\" & mstrUser & "\界面设置\" & App.EXEName & "\" & Me.Name, "WinLeft", Screen.Width / 4))
    Me.Top = NVL(GetSetting("ZLSOFT", "私有模块\" & mstrUser & "\界面设置\" & App.EXEName & "\" & Me.Name, "Wintop", Screen.Height / 4))
    Me.Width = NVL(GetSetting("ZLSOFT", "私有模块\" & mstrUser & "\界面设置\" & App.EXEName & "\" & Me.Name, "WinWidth", Screen.Width / 2))
    Me.Height = NVL(GetSetting("ZLSOFT", "私有模块\" & mstrUser & "\界面设置\" & App.EXEName & "\" & Me.Name, "WinHeight", Screen.Height / 2))
    ucPage.PageRecord = NVL(GetSetting("ZLSOFT", "私有模块\" & mstrUser & "\界面设置\" & App.EXEName & "\" & Me.Name, "PageRecord", 6))

    picImage.Width = NVL(GetSetting("ZLSOFT", "私有模块\" & mstrUser & "\界面设置\" & App.EXEName & "\" & Me.Name, "MiniatureW", 3000))
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
    
    With Me.cbrMain.Options
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
        
        If mblnIsMark Then
            Set cbrControl = .Add(IIf(mblnIsMark, xtpControlButton, xtpControlSplitButtonPopup), conMenu_Process_Save, "保存"): cbrControl.ToolTipText = "保存"
        Else
            Set cbrControl = .Add(xtpControlButton, conMenu_Process_SaveToStady, "存为检查图"): cbrControl.ToolTipText = "保存到检查图像"
            If Not mblnDoShiled Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Process_SaveToReport, "存为报告图"): cbrControl.ToolTipText = "保存到报告图像"
            End If
        End If
        
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Window, "亮度"): cbrControl.ToolTipText = "调节亮度/对比度": cbrControl.Visible = Not mblnIsMark
        cbrControl.Checked = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Zoom, "缩放"): cbrControl.ToolTipText = "缩放图像": cbrControl.Visible = Not mblnIsMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RectZoom, "裁剪"): cbrControl.ToolTipText = "裁剪采集图像": cbrControl.IconId = 3201: cbrControl.Visible = Not mblnIsMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RRotate, "顺时"): cbrControl.ToolTipText = "顺时针旋转": cbrControl.Visible = Not mblnIsMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_LRotate, "逆时"): cbrControl.ToolTipText = "逆时针旋转": cbrControl.Visible = Not mblnIsMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Sharpness, "锐化"): cbrControl.ToolTipText = "锐化": cbrControl.Visible = Not mblnIsMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Filter, "平滑"): cbrControl.ToolTipText = "平滑": cbrControl.Visible = Not mblnIsMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Line, "直线"): cbrControl.ToolTipText = "直线标注": cbrControl.Visible = Not mblnIsMark
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Arrow, "箭头"): cbrControl.ToolTipText = "箭头标注": cbrControl.Visible = Not mblnIsMark
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
    
    Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_Num * 100, "①"): objControl.ToolTipText = "自动递增数字编号": objControl.Category = 0: objControl.IconId = 0
    
    For i = 1 To 9
        Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_Num * 100 + i, i): objControl.ToolTipText = "数字编号" & i: objControl.Category = i: objControl.IconId = 0
    Next
    
End Sub

Private Sub LoadComText(mnuParent As Object)
    Dim objControl As CommandBarControl
    Dim arrTemp() As String
    Dim i As Long
    
    arrTemp = Split(mstrTemp, "|")
    
    For i = 0 To UBound(arrTemp)
        If Len(arrTemp(i)) > 0 Then
            Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_Word * 100 + i + 1, arrTemp(i)): objControl.ToolTipText = "文本标注": objControl.Category = i + 1: objControl.IconId = 0
        End If
    Next
End Sub


Private Sub lstMemoText_DblClick()
    cbxMemoText.Text = lstMemoText.List(lstMemoText.ListIndex)
    lstMemoText.Visible = False
    
    cbxMemoText.SelStart = 0
    cbxMemoText.SelLength = Len(cbxMemoText.Text)
    cbxMemoText.SetFocus
End Sub

Private Sub lstMemoText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        Call zlControl.CboSetText(cbxMemoText, lstMemoText.List(lstMemoText.ListIndex))
    End If
End Sub

Private Sub lstMemoText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlControl.CboSetText(cbxMemoText, lstMemoText.List(lstMemoText.ListIndex))
    End If
    
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then lstMemoText.Visible = False
End Sub

Private Sub picCboDropDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picCboDropDown.BorderStyle = 1
End Sub

Private Sub picCboDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picCboDropDown.BorderStyle = 0
End Sub

Private Sub picImage_Resize()
    
    Me.ucMiniature.Left = Me.picImage.Left
    Me.ucMiniature.Top = 0
    Me.ucMiniature.Width = Me.picImage.Width - 50
    Me.ucMiniature.Height = Abs(Me.picImage.Height - Me.ucPage.Height)

    Me.ucPage.Left = Me.picImage.Left
    Me.ucPage.Width = Me.picImage.Width - 50
    Me.ucPage.Top = Me.ucMiniature.Top + Me.ucMiniature.Height
End Sub

Private Sub picMemo_Resize()
    '摆放备注文字
    Me.lblMemoText.Left = 100
    Me.lblMemoText.Top = 200

    Me.cbxMemoText.Left = Me.lblMemoText.Left + Me.lblMemoText.Width
    Me.cbxMemoText.Top = Me.lblMemoText.Top - 100
    Me.cbxMemoText.Width = Abs(Me.ScaleWidth - Me.cbxMemoText.Left - 250 - cmdInsert.Width - cmdFont.Width - cmdAdd.Width)
    
    Me.picCboDropDown.Left = Me.cbxMemoText.Left + Me.cbxMemoText.Width - 270
    Me.picCboDropDown.Top = Me.cbxMemoText.Top + 30
    
    Me.cmdAdd.Left = Me.cbxMemoText.Left + Me.cbxMemoText.Width
    Me.cmdAdd.Top = Me.cbxMemoText.Top
    
    Me.cmdInsert.Left = Me.cmdAdd.Left + Me.cmdAdd.Width
    Me.cmdInsert.Top = Me.cbxMemoText.Top

    Me.cmdFont.Left = Me.cmdInsert.Left + Me.cmdInsert.Width
    Me.cmdFont.Top = Me.cmdInsert.Top
End Sub

Private Sub Timer1_Timer()
    Dim ptWin As POINTAPI

    GetCursorPos ptWin

    If mlngState = 1 Then
        If ptWin.X >= Me.Left / 15 And ptWin.X <= (Me.Left + Me.Width) / 15 And ptWin.Y >= Me.Top / 15 And ptWin.Y <= (Me.Top + Me.Height) / 15 Then
            mblnPreView = False
            Call RefrshObjVisible
            Call RefreshFace
            
            If DViewer.Images.Count > 0 Then
                Call ClearHint(DViewer.Images(1))
            End If
            
'            Me.cbrMain.FindControl(, conMenu_Process_Window).Parent.Visible = True
'            Timer4.Enabled = True
            mlngState = 2
            
            
            Timer2.Enabled = False
        End If
    ElseIf mlngState = 2 Then
        GetCursorPos ptWin

        If ptWin.X < Me.Left / 15 Or ptWin.X > (Me.Left + Me.Width) / 15 Or ptWin.Y < Me.Top / 15 Or ptWin.Y > (Me.Top + Me.Height) / 15 Then
            mblnPreView = True
            Call RefrshObjVisible
            Call RefreshFace
            
            If DViewer.Images.Count > 0 Then
                Call DrawHintTag(DViewer.Images(1))
            End If
'            Timer3.Enabled = True
            Me.cbrMain.FindControl(, conMenu_Process_Window).Parent.Visible = False
            If mlngPreViewTime > 0 And mlngWinType = 1 Then
                Timer2.Enabled = True
            End If
            mlngState = 1
        End If
    Else
        Timer1.Enabled = False
        Timer2.Enabled = False
    End If
End Sub



Private Sub Timer2_Timer()
    If mlngWinType = 1 And mlngPreViewTime > 0 Then
        Call UnloadMe
    End If
End Sub

'Private Sub Timer3_Timer()
'    Dim lngWidth As Long
'
'    If picMemo.Top <= Me.ScaleHeight Then
'        picMemo.Top = picMemo.Top + 50
'        DViewer.Height = DViewer.Height + 50
'    End If
'
'    If ucSplitter.Left > -135 Then
'        lngWidth = IIf(ucSplitter.Left - 200 < -135, ucSplitter.Left, 200)
'        picImage.Left = picImage.Left - lngWidth
'        ucSplitter.Left = ucSplitter.Left - lngWidth
'    End If
'
'    If picMemo.Top >= Me.ScaleHeight And ucSplitter.Left <= -135 Then
'        Timer3.Enabled = False
'    End If
'End Sub
'
'Private Sub Timer4_Timer()
'    If picMemo.Top > Me.Height - 600 Then
'        picMemo.Top = picMemo.Top - 50
'        DViewer.Height = DViewer.Height - 50
'    End If
'
'    If picImage.Width > 0 Then
'
'    End If
'
'    If picMemo.Top <= Me.Height - 600 And picImage.Width <= 0 Then
'        Timer4.Enabled = False
'    End If
'End Sub

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
    If la.Tag <> "" And Not la.TagObject Is Nothing Then
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
    
'    If mintMouseState <> msFixText Or mintTextIndex = 11 Then
'        MsgBox "请先选择一个文字标注按钮，然后才能设置。", vbOKOnly, CON_STR_HINT_TITLE
'        Exit Sub
'    End If
    
    strTemp = InputBox("请输入新的文字标注配置，格式为“简码1=说明1|简码2=说明2|...”。", "文字标注设置", Replace(mstrTemp, "[+]", "|"))
    
    If strTemp = "" Then Exit Sub
    
    If InStr(strTemp, "=") = 0 Then
        MsgBox "输入的格式不正确，应该按照“简码=说明”方式输入，请检查后重新设置。", vbOKOnly, CON_STR_HINT_TITLE
        Exit Sub
    End If
     
    '输入成功，使用这个新的文字标注，同时保存到注册表中
    mstrTemp = strTemp

    cbrMain.FindControl(, conMenu_Process_TextTag).CommandBar.Controls.DeleteAll
    
    LoadComText cbrMain.FindControl(, conMenu_Process_TextTag)
    Call SaveSetting("ZLSOFT", "公共模块\zl9PACSWork\frmReportImageEdit", "简明文字标注", Replace(mstrTemp, "|", "[+]"))
    
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
    Dim strText() As String
    Dim i  As Integer
    
    On Error GoTo err
    
    mstrTemp = GetSetting("ZLSOFT", "公共模块\zl9PACSWork\frmReportImageEdit", "简明文字标注", G_STR_TAG)
    
    mstrTemp = Replace(mstrTemp, "[+]", "|")
'    If strTemp = "" Then
'        '使用默认值，不需要设置
'        Exit Sub
'    End If
    
'    strText = Split(strTemp, "[+]")
'    If UBound(strText) <> 10 Then
'        '数据不符合格式，使用默认值
'        Exit Sub
'    End If
'
'    For i = 0 To 10
'        cmdTextLabel(i).Caption = strText(i)
'    Next i
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subLabelCopyRebuild(Simg As DicomImage, oImg As DicomImage)
'------------------------------------------------
'功能：重建图像的标注关联关系
'参数：sImg--源图像；oImg--目标图像
'返回：无
'------------------------------------------------
    Dim l As DicomLabel
    For Each l In oImg.Labels
        If Not l.TagObject Is Nothing Then
            If Simg.Labels.IndexOf(l.TagObject) <> 0 Then
                Set l.TagObject = oImg.Labels(Simg.Labels.IndexOf(l.TagObject))
            End If
        End If
    Next
End Sub



Public Sub UnloadMe()
    Set mobjDownLoadImages = Nothing
    Set mobjService = Nothing
    
    Unload Me
End Sub

Public Function IsRefresh() As Boolean
    IsRefresh = mblnPreView
End Function

Private Sub ucMiniature_OnSelChange(ByVal lngOldIndex As Long, ByVal lngNewIndex As Long)
    On Error GoTo err
    Dim dcmImage As DicomImage
    Dim i As Long
    
    If mblnIsMark Then Exit Sub
    
    '切换图像，将处理缓存到缩略图中，并记录修改状态
    '缩放和拖动图像不保存
    If lngOldIndex > 0 Then
        ucMiniature.ImgViewer.Images(lngOldIndex).Labels.Clear
        
        If DViewer.Images.Count > 0 Then
            '裁剪
            If mblnCase Then
                DViewer.Images(1).Copy
                ucMiniature.ImgViewer.Images(lngOldIndex).Paste
            Else
                Call CopyImages(DViewer.Images(1), ucMiniature.ImgViewer.Images(lngOldIndex))
            End If
            
            If mblnIsChanged Then
                ucMiniature.ImgViewer.Images(lngOldIndex).Tag.IsChanged = True
            End If
        End If
    End If
    
    DViewer.Images.Clear
    
    Set dcmImage = ucMiniature.ImgViewer.Images(lngNewIndex)
    DViewer.Images.Add dcmImage
    
    
    
    If DViewer.Images.Count > 0 Then
        ClearLable DViewer.Images(1)
        DViewer.Refresh
        
        Set mOldImage = dcmImage
    Else
        Set mOldImage = Nothing
    End If
    
    mblnIsChanged = False
    mblnCase = False
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ucPage_OnBeforeImageChange(Cancel As Boolean)
    If IsChanged Then
        If MsgBox("该操作将清除尚未保存的处理，是否继续？", vbYesNo, Me.Caption) = vbNo Then
            Cancel = True
        Else
            Call InitChangedState
        End If
    End If
End Sub

Private Sub ucPage_OnItemChange(ByVal lngPageIndex As Long, ByVal lngPageRecord As Long)
    ucMiniature.UpdateSelectIndex lngPageIndex
End Sub

Private Sub ucPage_OnPageChange(ByVal lngPageIndex As Long, ByVal lngPageCount As Long)
    Dim arrImages() As String
    
    Call mobjDownLoadImages.DownloadImages(arrImages, mstrQueryValue, (lngPageIndex - 1) * lngPageCount + 1, lngPageIndex * lngPageCount, False, mblnMoved)
    
    ucMiniature.RefreshImage arrImages
    
    ucMiniature.UpdateSelectIndex 1
End Sub

Private Sub ucPage_OnPageRecordChange(ByVal lngPageRecord As Long)
    Dim arrImages() As String
    
    Call mobjDownLoadImages.DownloadImages(arrImages, mstrQueryValue, ucPage.PageIndex * lngPageRecord + 1, (ucPage.PageIndex + 1) * lngPageRecord, False, mblnMoved)
    ucMiniature.RefreshImage arrImages
    
    ucMiniature.UpdateSelectIndex 1
End Sub

Public Sub AutoUnload()
    If mblnIsUnloud Then
        Timer2.Enabled = True
    End If
End Sub

Private Sub DrawHintTag(dcmImg As DicomImage)
    Dim lRpt As DicomLabel
    Dim i As Integer
     
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
        .Font.Size = 20
        .Font.Bold = False
        .ForeColour = &HCBBECB
        .Left = 120
        .Top = 20
        .Text = "...更多操作请点击..."
        .Shadow = doShadowBottomRight
        .Alignment = doAlignCentre
        .Visible = True
        .Tag = "HINT"
    End With
    
    dcmImg.Labels.Add lRpt
    
    dcmImg.Refresh False
End Sub

Private Sub ClearHint(dcmImage As DicomImage)
    Dim i As Long
    
    For i = 1 To dcmImage.Labels.Count
        If dcmImage.Labels(i).Tag = "HINT" Then
            dcmImage.Labels.Remove i
            Exit For
        End If
    Next
    
    dcmImage.Refresh False
End Sub

Public Sub AfterSaveStudy(dcmImage As DicomImage)
    
    If ucMiniature.ImgViewer.Images.Count < ucPage.PageRecord Then
        ucMiniature.AddImage dcmImage
    Else
        ucPage.RecordCount = ucPage.RecordCount + 1
    End If
    
    RaiseEvent AfterSaveStady
End Sub

Private Sub ucSplitter_OnMoveEnd()
    If lstMemoText.Visible Then lstMemoText.ZOrder
End Sub

Private Function IsChanged() As Boolean
'是否有处理过尚未保存的图像
    Dim i As Long
    
    IsChanged = False
    
    If Not mblnIsMark Then
        If ucMiniature.ImgViewer.Images.Count < 1 Then Exit Function
        
        If mblnIsChanged Then
            IsChanged = True
            Exit Function
        End If
        For i = 1 To ucMiniature.ImgViewer.Images.Count
            If ucMiniature.ImgViewer.Images(i).Tag.IsChanged Then
                IsChanged = True
                Exit Function
            End If
        Next
    Else
        If mblnIsChanged Then
            IsChanged = True
            Exit Function
        End If
    End If
    
End Function

Private Sub InitChangedState()
    Dim i As Long
    
    If ucMiniature.ImgViewer.Images.Count < 1 Then Exit Sub
    
    For i = 1 To ucMiniature.ImgViewer.Images.Count
        ucMiniature.ImgViewer.Images(i).Tag.IsChanged = False
    Next
End Sub

Private Sub CopyImages(dcmImage As DicomImage, dcmSub As DicomImage)
'将图像的处理临时缓存到缩略图中
    Dim i As Long
    
    If mblnCase Then Exit Sub
    
    dcmSub.Zoom = dcmImage.Zoom
    dcmSub.StretchToFit = dcmImage.StretchToFit
    dcmSub.ScrollX = dcmImage.ScrollX
    dcmSub.ScrollY = dcmImage.ScrollY
    dcmSub.Width = dcmImage.Width
    dcmSub.Level = dcmImage.Level
    
    dcmSub.UnsharpEnhancement = dcmImage.UnsharpEnhancement
    dcmSub.FilterLength = dcmImage.FilterLength
    dcmSub.RotateState = dcmImage.RotateState
    
    For i = 1 To dcmImage.Labels.Count
        If dcmImage.Labels(i).Tag <> "HINT" Then
            dcmSub.Labels.Add dcmImage.Labels(i)
        End If
    Next

    dcmSub.Refresh False
End Sub

