VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.UserControl ucImagePreview 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7605
   ScaleHeight     =   3795
   ScaleWidth      =   7605
   ToolboxBitmap   =   "ucImagePreview.ctx":0000
   Begin zl9PACSWork.ucSplitPage ucPage 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   6210
      _ExtentX        =   10504
      _ExtentY        =   582
      PageCount       =   0
      PageRecord      =   9
   End
   Begin DicomObjects.DicomViewer dcmMiniImage 
      Height          =   3135
      Left            =   120
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
   End
End
Attribute VB_Name = "ucImagePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const M_STR_SELECT_TAG As String = "SELECT"
Private Const M_STR_BORDER_TAG As String = "BORDER"

Public Enum TQueryLevel
    slStudy = 0     '检查
    slSeries = 1    '序列
    slImage = 2     '图像
End Enum


Private mstrQueryValue As String         '检查医嘱ID
Private mblnMoved As Boolean             '数据是否被转存
Private mslQueryLevel As TQueryLevel      '图像显示级别
Private mblnQueryTmpRecord As Boolean

Private mcnOracle As ADODB.Connection   '数据库连接对象
Private mblnIsShowCheckbox As Boolean   '是否显示勾选框
Private mblnEnable As Boolean           '是否可进行编辑
Private mlngMouseMoveZoom As Double     '鼠标在图像上移动时，显示大图的放大倍数，如果为0则不显示大图
Private mblnIsShowPopup As Boolean      '是否显示右键菜单
Private mblnIsAutoHidePageControl As Boolean


Private mMultiCols As Long
Private mMultiRows As Long

Private mlngSelectIndex As Long


Public Event OnSelChange(ByVal lngSelectedIndex As Long)
Public Event OnCheckChange(ByVal lngSelectedIndex As Long, ByVal blnSelected As Boolean)
Public Event OnClick(ByVal lngSelectedIndex As Long)
Public Event OnDbClick(ByVal lngSelectedIndex As Long, ByRef blnContinue As Boolean)


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


'是否显示图像勾选框
Property Get ShowCheckBox() As Long
    ShowCheckBox = mblnIsShowCheckbox
End Property

Property Let ShowCheckBox(value As Long)
    mblnIsShowCheckbox = value
End Property

'鼠标移动到图像上的放大倍数，如果为0则不进行放大
Property Get MouseMoveZoom() As Long
    MouseMoveZoom = mlngMouseMoveZoom
End Property

Property Let MouseMoveZoom(value As Long)
    mlngMouseMoveZoom = value
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
Property Get ImgViewer() As DicomViewer
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
            ImgChecked = Not objLabs(i).Transparent
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


'初始化图像预览
Public Sub InitImgPreview(cnOracle As ADODB.Connection)
    Set mcnOracle = cnOracle
    
    Call dcmMiniImage.Images.Clear
End Sub


Public Sub RefreshImage(ByVal slQueryLevel As TQueryLevel, ByVal strQueryValue As String, _
    ByVal blnMoved As Boolean, Optional ByVal blnFoceRefresh As Boolean = False, Optional ByVal blnTmpRecord As Boolean = False)
'刷新图像显示
    Dim rsData As ADODB.Recordset
    
    If mstrQueryValue = strQueryValue And Not blnFoceRefresh Then Exit Sub
    
    mstrQueryValue = strQueryValue
    mslQueryLevel = slQueryLevel
    mblnQueryTmpRecord = blnTmpRecord
    mblnMoved = blnMoved
    
    ucPage.RecordCount = 0
    mlngSelectIndex = 0
    
    Call RefreshPageControl
    
    '清除图像
    Call ClearCurrentPageImage
    
    If strQueryValue = "" Then Exit Sub
    
    '配置分页组件
    Call ConfigPageControl(slQueryLevel, strQueryValue, blnTmpRecord)
    
    '读取图像数据
    Set rsData = GetImageViewData(slQueryLevel, strQueryValue, 1, ucPage.PageRecord, blnTmpRecord)
    
    '加载图像信息
    Call LoadViewImageToFace(rsData, dcmMiniImage)
    
    If mblnIsShowCheckbox Then
        '绘制选择框
        Call DrawImageSelectBorder(dcmMiniImage)
    End If
End Sub


Private Sub RefreshPageControl()
'刷新分页组件显示
On Error Resume Next
    If Not mblnIsAutoHidePageControl Then Exit Sub
    
    ucPage.Visible = IIf(ucPage.RecordCount <= ucPage.PageRecord, False, True)
    
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
    If ErrCenter() = 1 Then Resume
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


Private Sub LoadViewImageToFace(rsCurImageData As ADODB.Recordset, dcmViewer As DicomViewer)
'加载预览图像到界面
    Dim strTmpFile As String
    Dim strCachePath As String
    
    Dim curImage As DicomImage
    
    Dim objFile As New Scripting.FileSystemObject
    
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    
    Dim dcmTag As clsImageTagInf
    
    
    mlngSelectIndex = 0
    
    dcmViewer.Images.Clear
    
    If rsCurImageData.RecordCount > 0 Then

        Call ConfigImgDisplayFormat(IIf(rsCurImageData.RecordCount < ucPage.PageRecord, rsCurImageData.RecordCount, ucPage.PageRecord))
        
        '创建本地目录
        strCachePath = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
        MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsCurImageData("URL")))
        
        Do While Not rsCurImageData.EOF
            '循环加载图像到DicomViewer中
            strTmpFile = strCachePath & Nvl(rsCurImageData("URL"))
            
            '设置音视频的显示文件，如果为音视频文件时，该过程将不从服务器中直接下载数据文件
            If Nvl(rsCurImageData!动态图, IMGTAG) = VIDEOTAG Then
                strTmpFile = IIf(Len(App.Path) > 3, App.Path & "\..\附加文件\Avi.bmp", App.Path & "..\附加文件\Avi.bmp")
            ElseIf Nvl(rsCurImageData!动态图, IMGTAG) = AUDIOTAG Then
                strTmpFile = IIf(Len(App.Path) > 3, App.Path & "\..\附加文件\wav.bmp", App.Path & "..\附加文件\wav.bmp")
            End If
            
            If Dir(strTmpFile) = vbNullString Then
                '本地缓存图像不存在，则读取FTP图像
                
                '建立FTP连接
                If Nvl(rsCurImageData("设备号1")) <> vbNullString And Inet1.hConnection = 0 Then
                    If Inet1.FuncFtpConnect(Nvl(rsCurImageData("Host1")), Nvl(rsCurImageData("User1")), Nvl(rsCurImageData("Pwd1"))) = 0 Then
                        If Nvl(rsCurImageData("设备号2")) <> vbNullString Then
                            If Inet2.FuncFtpConnect(Nvl(rsCurImageData("Host2")), Nvl(rsCurImageData("User2")), Nvl(rsCurImageData("Pwd2"))) = 0 Then
                                MsgBoxD Parent, "FTP不能正常连接，请检查网络设置。"
                                Exit Sub
                            End If
                        Else
                            MsgBoxD Parent, "FTP不能正常连接，请检查网络设置。"
                            Exit Sub
                        End If
                    End If
                End If
                
                If Inet1.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsCurImageData("Root1")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL"))) <> 0 Then
                    '从设备号1提取图像失败，则从设备号2提取图像
                    If Nvl(rsCurImageData("设备号2")) <> vbNullString Then
                        If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect Nvl(rsCurImageData("Host2")), Nvl(rsCurImageData("User2")), Nvl(rsCurImageData("Pwd2"))
                        Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsCurImageData("Root2")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL")))
                    End If
                End If
            End If
  
            If Dir(strTmpFile) <> vbNullString Then
               If Nvl(rsCurImageData!动态图, IMGTAG) <> VIDEOTAG And Nvl(rsCurImageData("动态图"), IMGTAG) <> AUDIOTAG Then
                    Set curImage = dcmViewer.Images.ReadFile(strTmpFile)
                    
                    '设置图像标记
                    Set dcmTag = New clsImageTagInf
                    dcmTag.Tag = Nvl(rsCurImageData!动态图, IMGTAG)
                                        
                    Set curImage.Tag = dcmTag
                    
                    With curImage
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                Else
                    Set curImage = New DicomImage
                    
                    On Error GoTo continue
                        Call curImage.FileImport(strTmpFile, "DIB/BMP")
continue:
                        Set dcmTag = New clsImageTagInf
                        
                        dcmTag.Tag = Nvl(rsCurImageData!动态图, VIDEOTAG)
                        dcmTag.EncoderName = Nvl(rsCurImageData("编码名称"), "")
                        dcmTag.CaptureTime = Nvl(rsCurImageData("采集时间"))
                        
                        If Nvl(rsCurImageData("动态图"), VIDEOTAG) = VIDEOTAG Then
                            dcmTag.VideoFile = strCachePath & Nvl(rsCurImageData("URL")) & ".avi"
                        Else
                            dcmTag.VideoFile = strCachePath & Nvl(rsCurImageData("URL")) & ".wav"
                        End If
                        
                        dcmTag.RecordTimeLen = Val(Nvl(rsCurImageData("录制长度"), "0"))
                        
'                        '如果是视频录像文件，则在播放时进行下载
'                        If Trim(dcmTag.VideoFile) <> "" And Dir(dcmTag.VideoFile) <> "" Then
'                            Name dcmTag.VideoFile As dcmTag.VideoFile & ".avi"
'                        End If
                        
                        Set curImage.Tag = dcmTag
                        
                        curImage.InstanceUID = Nvl(rsCurImageData("图像UID"))
                        curImage.SeriesUID = Nvl(rsCurImageData("序列UID"))
                        curImage.StudyUID = Nvl(rsCurImageData("检查UID"))
                        
                    
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
        
        Inet1.FuncFtpDisConnect
        Inet2.FuncFtpDisConnect
    Else
        dcmViewer.MultiColumns = 1
        dcmViewer.MultiRows = 1
    End If
End Sub


Public Sub PlayMedia(ByVal lngMediaIndex As Long)
'播放指定索引处的媒体

End Sub


Private Sub ConfigPageControl(ByVal slQueryLevel As TQueryLevel, ByVal strSearchValue As String, ByVal blnTmpRecord As Boolean)
'配置分页控件
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    
    
    Select Case slQueryLevel
        Case slStudy
            strSql = "select count(1)  as 返回值 from 影像检查图象 a, 影像检查序列 b where a.序列UID=b.序列UID and b.检查UID=[1]"
        Case slSeries
            strSql = "select count(1)  as 返回值 from 影像检查图象  where  序列UID=[1]"
        Case slImage
            strSql = "select count(1)  as 返回值 from 影像检查图象  where  图像UID=[1]"
    End Select
    
    '如果查询临时记录，则需要将查询表替换为临时存储数据的表
    If blnTmpRecord Then
        strSql = Replace(strSql, "影像检查", "影像临时")
    Else
        If mblnMoved Then
            strSql = Replace(strSql, "影像检查图象", "H影像检查图象")
            strSql = Replace(strSql, "影像检查序列", "H影像检查序列")
        End If
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询图像数量", strSearchValue)
    
    If rsData.RecordCount > 0 Then
        lngRecordCount = Nvl(rsData!返回值)
    Else
        lngRecordCount = 0
    End If

    
'    ucPage.PageRecord = mlngPageRecord
    ucPage.RecordCount = lngRecordCount
    
    Call RefreshPageControl
End Sub



Private Function GetImageViewData(ByVal slQueryLevel As TQueryLevel, ByVal strSearchValue As String, _
    ByVal lngCurPage As Long, ByVal lngPageRecord As Long, ByVal blnTmpRecord As Boolean) As ADODB.Recordset
'获取预览图像数据
'intSearchType:0-按检查uid搜索,1-按序列UID搜索,2-按图像UID搜索

    Dim strSql As String
    Dim lngStartRecord As Long
    Dim lngEndRecord As Long
    

        
    strSql = "Select rownum as 顺序号, A.图像号,D.FTP用户名 As User1,D.FTP密码 As Pwd1," & _
        "D.IP地址 As Host1,'/'||D.Ftp目录||'/' As Root1," & _
        "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/'||A.图像UID As URL,d.设备号 as 设备号1, " & _
        "E.FTP用户名 As User2,E.FTP密码 As Pwd2," & _
        "E.IP地址 As Host2,'/'||E.Ftp目录||'/' As Root2," & _
        "e.设备号 as 设备号2, A.图像UID,C.检查UID,B.序列UID,A.动态图,A.编码名称,A.采集时间, A.录制长度 " & _
        "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) "
        
    If blnTmpRecord Then
        strSql = Replace(strSql, "影像检查", "影像临时")
    Else
        If mblnMoved Then
            strSql = Replace(strSql, "影像检查图象", "H影像检查图象")
            strSql = Replace(strSql, "影像检查序列", "H影像检查序列")
            strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
        End If
    End If

    Select Case slQueryLevel
        Case slStudy
            strSql = strSql & " and C.检查UID=[1]"
        Case slSeries
            strSql = strSql & " and B.序列UID=[1]"
        Case slImage
            strSql = strSql & " and A.图像UID=[1]"
    End Select
    
    lngStartRecord = (lngCurPage - 1) * lngPageRecord + 1
    lngEndRecord = lngCurPage * lngPageRecord
    
    strSql = "select /*+RULE*/ * from (" & strSql & " order by b.序列UID, a.图像号) where 顺序号>=" & lngStartRecord & " and 顺序号<=" & lngEndRecord
    
    Set GetImageViewData = zlDatabase.OpenSQLRecord(strSql, "查询图像信息", strSearchValue)
End Function

Public Sub AddImage(img As DicomImage, Optional objImgTag As clsImageTagInf = Nothing)
'增加图像
    Dim i As Long
    
    If dcmMiniImage.Images.Count < ucPage.PageRecord Then
        Call ConfigImgDisplayFormat(dcmMiniImage.Images.Count + 1)
        
        Call dcmMiniImage.Images.Add(img)
    Else
        '移动图像
        For i = 2 To dcmMiniImage.Images.Count
            Call dcmMiniImage.Images.Move(i, i - 1)
            dcmMiniImage.Images(i - 1).BorderColour = vbWhite
        Next i
        
        Call dcmMiniImage.Images.Remove(dcmMiniImage.Images.Count)
        dcmMiniImage.Images.Add img
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
    
    Call UpdateSelectIndex(dcmMiniImage.Images.Count)
    Call UpdateImageCount(1)
End Sub


Private Sub ShowAVInf(img As DicomImage, objImgTag As clsImageTagInf)
'显示音视频信息
    If objImgTag.Tag = VIDEOTAG Or objImgTag.Tag = AUDIOTAG Then
        Call AddVideoLabelToDicomImage(img, _
        IIf(objImgTag.Tag = VIDEOTAG, "录像时间：", "录音时间：") & objImgTag.CaptureTime, _
        IIf(objImgTag.Tag = VIDEOTAG, "录像长度：", "录音长度：") & objImgTag.RecordTimeLen & " 秒", _
        "编码名称：" & objImgTag.EncoderName)
    End If
End Sub

Public Sub DeleteImage(ByVal lngImgIndex As Long)
'删除图像
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
        '当图像页数改变时，重新刷新当前页图像显示
        Call ucPage.MovePage(ucPage.PageNumber)
        If dcmMiniImage.Images.Count > 0 Then Call UpdateSelectIndex(1)
    End If
    
'    If dcmMiniImage.Images.Count <= 0 Then
'        Call ucPage.MovePage(ucPage.PageNumber)
'
'        If dcmMiniImage.Images.Count > 0 Then Call UpdateSelectIndex(1)
'    End If
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
    If mlngSelectIndex <= 0 Or mlngSelectIndex > dcmMiniImage.Images.Count Then Exit Sub
    
    Call DoOnClick(mlngSelectIndex)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub dcmMiniImage_DblClick()
On Error GoTo errHandle
    Dim blnContinue As Boolean
    
    If dcmMiniImage.Images.Count <= 0 Then Exit Sub
    If mlngSelectIndex <= 0 Then Exit Sub

    blnContinue = True
    
    Call DoOnDbClick(mlngSelectIndex, blnContinue)
    
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
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub dcmMiniImage_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error Resume Next
    Dim i As Long
    Dim objLabs As DicomLabels
    Dim lngImgIndex As Long
    
    lngImgIndex = dcmMiniImage.ImageIndex(X, Y)
    
    Call UpdateSelectIndex(lngImgIndex)
    
    If mlngSelectIndex > 0 And mlngSelectIndex <= dcmMiniImage.Images.Count Then
        
        If mblnEnable Then
        
            '设置选择框状态
            Set objLabs = dcmMiniImage.LabelHits(X, Y, False, True, True)
            For i = 1 To objLabs.Count
                If objLabs(i).Tag = M_STR_SELECT_TAG Then
                    objLabs(i).Transparent = Not objLabs(i).Transparent
                    
                    '触发图像勾选事件
                    Call DoOnCheckChange(mlngSelectIndex, Not objLabs(i).Transparent)
                    
                    Exit For
                End If
            Next i
            
        End If
    End If
End Sub


Private Sub dcmMiniImage_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errHandle
    Dim blnShowImg As Boolean
    Dim intCurrImg As Integer
    
    '没有放大倍数或图像，则不进行图像缩放
    If mlngMouseMoveZoom = 0 Then Exit Sub
    If dcmMiniImage.Images.Count <= 0 Then Exit Sub
    
    '判断是否需要显示图像
    If (0 <= X * Screen.TwipsPerPixelX) And (X * Screen.TwipsPerPixelX <= dcmMiniImage.Width) And _
       (0 <= Y * Screen.TwipsPerPixelY) And (Y * Screen.TwipsPerPixelY <= dcmMiniImage.Height) Then
        blnShowImg = True
    End If
    
    If blnShowImg Then      '显示图像
        SetCapture dcmMiniImage.hWnd    '锁定鼠标
        
        intCurrImg = dcmMiniImage.ImageIndex(X, Y)
        
        If intCurrImg <> 0 Then
            '加载图像并显示
            frmShowImg.ShowMe dcmMiniImage.Images(intCurrImg), Me, 1, 0, 0, mlngMouseMoveZoom
        Else
            frmShowImg.HideMe
        End If
    Else        '关闭图像显示
        ReleaseCapture      '解锁鼠标
        frmShowImg.HideMe
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub dcmMiniImage_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errHandle
    Dim curPointer As PointAPI
    
    If Button = 2 And Not ucPage.Visible Then
        Call GetCursorPos(curPointer)
        
        Call ScreenToClient(hWnd, curPointer)  'ScreenToClient方法使用的单位为像素值
        Call PopupMenu(menuPopup, 0, ScaleX(curPointer.X, vbPixels, vbTwips), ScaleY(curPointer.Y, vbPixels, vbTwips))
        
    End If
errHandle:
End Sub

Private Sub dcmMiniImage_MouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)
On Error GoTo errHandle
    If Delta > 0 Then
        Call ucPage.LastPage
    Else
        Call ucPage.NextPage
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub mnuSplitPageTool_Click()
'显示分页工具栏
On Error GoTo errHandle
    ucPage.Visible = True
    
    Call UserControl_Resize
errHandle:
End Sub

Private Sub ucPage_OnPageChange(ByVal lngPageIndex As Long, ByVal lngPageCount As Long)
On Error GoTo errHandle
    Dim rsData As ADODB.Recordset
    
    Set rsData = GetImageViewData(mslQueryLevel, mstrQueryValue, lngPageIndex, lngPageCount, mblnQueryTmpRecord)
    Call LoadViewImageToFace(rsData, dcmMiniImage)
    
    
    If mblnIsShowCheckbox Then
        Call DrawImageSelectBorder(dcmMiniImage)
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub UserControl_Initialize()
    Set mcnOracle = Nothing
    
    dcmMiniImage.CellSpacing = 3
    mblnEnable = True
    
    mblnIsShowCheckbox = False
    mblnIsShowPopup = False
    
    mlngMouseMoveZoom = 0
    
    mstrQueryValue = ""
    mlngSelectIndex = 0
    
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
On Error Resume Next
    
    dcmMiniImage.Left = 0
    dcmMiniImage.Top = 0
    dcmMiniImage.Width = UserControl.ScaleWidth
    dcmMiniImage.Height = UserControl.ScaleHeight - IIf(ucPage.Visible, ucPage.Height + 60, 0)
    
    ucPage.Left = 0
    ucPage.Top = dcmMiniImage.Height + 30
    
    err.Clear
End Sub



Private Function GetImageRow(ByVal lngImageIndex As Long) As Integer
'取得当前所在行
    GetImageRow = CInt(lngImageIndex / dcmMiniImage.MultiColumns) + 1
End Function

Private Sub UserControl_Show()
    Call UserControl_Resize
End Sub



Private Sub UserControl_Terminate()
    Set mcnOracle = Nothing
End Sub




Private Sub DrawImageSelectBorder(dcmViewer As DicomViewer)
'绘制图像选择边框
    Dim i As Long
    
    Dim lSelect As DicomLabel
    Dim lBorder As DicomLabel

    
    '循环每一个图像，画标注
    For i = 1 To dcmViewer.Images.Count
        Call dcmViewer.Images(i).Labels.Clear
        
        Set lBorder = New DicomLabel

        lBorder.LabelType = 2            '边框
        lBorder.Width = 1000
        lBorder.Height = 1000
        lBorder.Left = 0
        lBorder.Top = 0
        lBorder.LineWidth = 2


        lBorder.ForeColour = vbYellow
        lBorder.BackColour = vbYellow


        lBorder.Transparent = True
        lBorder.ScaleWithCell = True
        lBorder.Tag = M_STR_BORDER_TAG

        lBorder.Visible = True
        dcmViewer.Images(i).Labels.Add lBorder
        

    
    
        Set lSelect = New DicomLabel
        
        lSelect.LabelType = 2            '矩形
        lSelect.Width = 18
        lSelect.Height = 18
        lSelect.Left = 1
        lSelect.Top = 1
        lSelect.LineWidth = 2
        
        lSelect.ForeColour = vbYellow
        lSelect.BackColour = vbRed
        
                
        lSelect.Transparent = True
        lSelect.ScaleWithCell = False
        lSelect.ImageTied = False

        lSelect.Tag = M_STR_SELECT_TAG
        
        lSelect.Visible = True
        dcmViewer.Images(i).Labels.Add lSelect
        
        dcmViewer.Images(1).BorderStyle = vbRed
    Next i
End Sub





Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    
    dcmMiniImage.CellSpacing = PropBag.ReadProperty("CellSpacing", 3)
    dcmMiniImage.BackColour = PropBag.ReadProperty("BackColor", vbBlack)
    mblnEnable = PropBag.ReadProperty("Enable", True)
    mblnIsShowCheckbox = PropBag.ReadProperty("ShowCheckbox", False)
    mblnIsShowPopup = PropBag.ReadProperty("ShowPopup", False)
    mlngMouseMoveZoom = PropBag.ReadProperty("MouseMoveZoom", 0)
    ucPage.PageRecord = PropBag.ReadProperty("PageImgCount", 5)
    
    err.Clear
End Sub



Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next

    Call PropBag.WriteProperty("CellSpacing", dcmMiniImage.CellSpacing, 3)
    Call PropBag.WriteProperty("BackColor", dcmMiniImage.BackColour, vbBlack)
    Call PropBag.WriteProperty("Enable", mblnEnable, True)
    Call PropBag.WriteProperty("ShowCheckbox", mblnIsShowCheckbox, False)
    Call PropBag.WriteProperty("ShowPopup", mblnIsShowPopup, False)
    Call PropBag.WriteProperty("MouseMoveZoom", mlngMouseMoveZoom, 0)
    Call PropBag.WriteProperty("PageImgCount", ucPage.PageRecord, 5)
    
    err.Clear
End Sub
