VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.UserControl ucCacheImages 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   ScaleHeight     =   8655
   ScaleWidth      =   5385
   Begin VB.VScrollBar vscImages 
      Height          =   8175
      Left            =   5040
      TabIndex        =   3
      Top             =   0
      Width           =   275
   End
   Begin VB.ComboBox cboCacheMark 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   8160
      Width           =   3375
   End
   Begin MSComCtl2.DTPicker dtpCacheDate 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   8160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   158793729
      CurrentDate     =   42674
   End
   Begin DicomObjects.DicomViewer dcmCacheImg 
      Height          =   7455
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4425
      _Version        =   262147
      _ExtentX        =   7805
      _ExtentY        =   13150
      _StockProps     =   35
      BackColor       =   0
      CellSpacing     =   2
   End
End
Attribute VB_Name = "ucCacheImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private mlngSelImgIndex As Long

Public Event OnDblClick()
Public Event OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)


Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

 

Property Get ImgCount() As Long
On Error GoTo errhandle
    ImgCount = dcmCacheImg.Images.Count
Exit Property
errhandle:
    ImgCount = 0
End Property


Public Sub SetFontSize(ByVal bytFontSize As Byte)
    FontSize = bytFontSize
    
    cboCacheMark.FontSize = bytFontSize
    Set dtpCacheDate.Font = Font
    
    dtpCacheDate.Height = cboCacheMark.Height
End Sub

Private Sub InitFace()
    dcmCacheImg.Move 0, 0, ScaleWidth - vscImages.Width, ScaleHeight - cboCacheMark.Height
    vscImages.Move ScaleWidth - vscImages.Width, 0, vscImages.Width, ScaleHeight - cboCacheMark.Height
    
    dtpCacheDate.Move 0, ScaleHeight - dtpCacheDate.Height, dtpCacheDate.Width, dtpCacheDate.Height
'    cboCacheMark.Move dtpCacheDate.Width, ScaleHeight - cboCacheMark.Height, ScaleWidth - vscImages.Width - dtpCacheDate.Width, cboCacheMark.Height
    cboCacheMark.Left = dtpCacheDate.Width
    cboCacheMark.Top = ScaleHeight - cboCacheMark.Height
    cboCacheMark.Width = ScaleWidth - dtpCacheDate.Width - vscImages.Width
     
End Sub

Public Sub Refresh()
    dtpCacheDate.value = Now
    Call LoadCacheMark(dtpCacheDate.value)
End Sub

Public Function IsSelected(Optional ByVal lngIndex As Long = 0) As Boolean
    Dim i As Long
    
    IsSelected = False
    If lngIndex = 0 Then
        For i = 1 To dcmCacheImg.Images.Count
            If dcmCacheImg.Images(i).BorderColour <> IMG_BACK_BORDER_COLOR Then
                IsSelected = True
                Exit Function
            End If
        Next
    Else
        IsSelected = IIf(dcmCacheImg.Images(lngIndex).BorderColour <> IMG_BACK_BORDER_COLOR, True, False)
    End If
End Function


Public Sub SyncAfterShow(objImg As Object, ByVal strAfterTag As String)
    Dim lngCurCount As Long
    
    If strAfterTag <> cboCacheMark.Text Then Exit Sub
    
    If vscImages.Max <= 0 Then
        lngCurCount = dcmCacheImg.Images.Count
        
        Call DrawBorder(objImg, 0)
        dcmCacheImg.Images.Add objImg
        
        If lngCurCount + 1 > dcmCacheImg.MultiColumns * dcmCacheImg.MultiRows Then
            vscImages.Min = 1
            vscImages.Max = 2
            vscImages.Enabled = True
        End If
    Else
        dcmCacheImg.Images.Add objImg
       vscImages.Max = vscImages.Max + 1
    End If
End Sub

Public Sub OpenCachePath()
'打开缓存路径
    Dim strCachePath As String
    
    strCachePath = GetCachePath(Format(dtpCacheDate.value, "YYYYMMDD"), cboCacheMark.Text)
    
    If DirExists(strCachePath) = False Then Call MkLocalDir(strCachePath)
    
    ShellExecute 0, "open", strCachePath, "", "", 1
End Sub

Public Function GetCachePath(ByVal strFmtDate As String, ByVal strMark As String) As String
    GetCachePath = FormatFilePath(SysRootPath & "\Apply\TmpAfterImage\" & strFmtDate & "\" & IIf(Len(strMark) <= 0, "", strMark & "\"))
End Function


Public Sub SelectedAll()
'全选
    Dim i As Long
    
    For i = 1 To dcmCacheImg.Images.Count
        Call DrawBorder(dcmCacheImg.Images(i), ColorConstants.vbRed, True)
    Next i
End Sub

Public Function GetSelects() As Long()
'获取选中的图像索引
'索引从1开始
    Dim i As Long
    Dim lngBound As Long
    Dim arySelIndex() As Long
    
    ReDim arySelIndex(0)
    
    For i = 1 To dcmCacheImg.Images.Count
        
        If IsSelected(i) Then
            '如果是非透明颜色,说明是被选中的图像
            lngBound = UBound(arySelIndex) + 1
            ReDim Preserve arySelIndex(lngBound)
            
            arySelIndex(lngBound) = i
        End If
    Next
    
    GetSelects = arySelIndex
End Function
 
Public Function GetImage(ByVal lngIndex As Long) As DicomImage
'获取图像
    Dim objSelImg As DicomImage
    
    Set GetImage = Nothing
    
    If lngIndex <= 0 Or lngIndex > dcmCacheImg.Images.Count Then Exit Function
    
    Set objSelImg = dcmCacheImg.Images(lngIndex)
    
    Set GetImage = objSelImg.SubImage(0, 0, objSelImg.SizeX, objSelImg.SizeY, 1, 1)
End Function

Public Sub DeleteCacheImg(Optional ByVal lngIndex As Long = 0)
'删除缓存图像
    Dim i As Long
    Dim strUid As String
    Dim objImg As DicomImage
    Dim strCachePath As String
    Dim strCacheFile As String
    
    strCachePath = GetCachePath(Format(dtpCacheDate.value, "yyyymmdd"), cboCacheMark.Text)
    
    If lngIndex = 0 Then
        Set objImg = dcmCacheImg.Images(mlngSelImgIndex)
        If objImg.BorderColour <> IMG_BACK_BORDER_COLOR Then
            strUid = objImg.InstanceUID
            
            If Len(strUid) > 0 Then
                strCacheFile = strCachePath & strUid
            Else
                strCacheFile = objImg.tag
            End If
            
            Set objImg = Nothing
            
            Call dcmCacheImg.Images.Remove(mlngSelImgIndex)
            Call RemoveFile(strCacheFile)
            
            vscImages.Max = vscImages.Max - 1
        End If
    Else
        For i = dcmCacheImg.Images.Count To 1 Step -1
            Set objImg = dcmCacheImg.Images(i)
            If objImg.BorderColour <> IMG_BACK_BORDER_COLOR Then
                
                strUid = objImg.InstanceUID
                
                If Len(strUid) > 0 Then
                    strCacheFile = strCachePath & strUid
                Else
                    strCacheFile = objImg.tag
                End If
                
                Set objImg = Nothing
                
                Call dcmCacheImg.Images.Remove(i)
                Call RemoveFile(strCacheFile)
                
                vscImages.Max = vscImages.Max - 1
            End If
        Next
        
        If vscImages.Max <= 0 Then
            vscImages.Max = 0
            vscImages.Min = 0
            vscImages.value = 0
        End If
    End If
    
    dcmCacheImg.Refresh
    
    If dcmCacheImg.Images.Count <= 0 Then
'        Call DeleteFolder(Replace(strCachePath & "\", "\\", ""))
        RmDir strCachePath
        Call cboCacheMark.RemoveItem(cboCacheMark.ListIndex)
    End If
End Sub


Private Sub LoadCacheMark(ByVal dtDate As Date)
'载入后台缓存标记
    Dim strCachePath As String
    Dim objFileSys As New FileSystemObject
    Dim objFolder As Folder
    Dim objSubFolder As Folder
    
    strCachePath = GetCachePath(Format(dtDate, "YYYYMMDD"), "")
    
    cboCacheMark.Clear
    dcmCacheImg.Images.Clear
    vscImages.Max = 0
    vscImages.Min = 0
    vscImages.Enabled = False
    
    If DirExists(strCachePath) = False Then Exit Sub
    Set objFolder = objFileSys.GetFolder(strCachePath)
     
    For Each objSubFolder In objFolder.SubFolders
        cboCacheMark.AddItem objSubFolder.Name
    Next
    
    If cboCacheMark.ListCount > 0 Then cboCacheMark.ListIndex = cboCacheMark.ListCount - 1
End Sub

Private Function GetPageRecordCount() As Long
    Dim dblMinArea As Double
    Dim dblViewerArea As Double
    
    dblMinArea = CDbl(2700) * CDbl(2800)
    dblViewerArea = dcmCacheImg.Width * dcmCacheImg.Height
    
    GetPageRecordCount = CInt(dblViewerArea / dblMinArea)
End Function

Private Sub LoadCacheImages(ByVal dtDate As Date, ByVal strMark As String)
'载入缓存图像
    Dim strMarkPath As String
    Dim objFileSys As New FileSystemObject
    Dim objFolder As Folder
    Dim objFile As File
    Dim objImg As DicomImage
    Dim objImgInfos As clsBgImgInfo
    
    strMarkPath = GetCachePath(Format(dtDate, "YYYYMMDD"), strMark)
    
    dcmCacheImg.Images.Clear
    
    If DirExists(strMarkPath) = False Then Exit Sub
    Set objFolder = objFileSys.GetFolder(strMarkPath)
    
    If objFolder.Files.Count > dcmCacheImg.MultiColumns * dcmCacheImg.MultiRows Then
        vscImages.Min = 1
        vscImages.Max = objFolder.Files.Count - (dcmCacheImg.MultiColumns * dcmCacheImg.MultiRows) + 1
        vscImages.value = 1
        
        vscImages.Enabled = True
    Else
        vscImages.Min = 0
        vscImages.Max = 0
        vscImages.Enabled = False
    End If
    
    
    For Each objFile In objFolder.Files
        Set objImg = dcmCacheImg.Images.ReadFile(objFile.Path)
        objImg.tag = objFile.Path
        
        If Len(objImg.InstanceUID) <= 0 Then
            Set objImgInfos = New clsBgImgInfo
            
            objImgInfos.LoadState = lsError
            objImgInfos.ErrorInfo = "未能识别的格式文件"
            objImgInfos.Redo = 0
            
            DrawErrorInfo objImg, objImgInfos
            
'            objImg.InstanceUID = objFile.Name
            
        End If
        
        Call DrawBorder(objImg, 0)
    Next
    
    If dcmCacheImg.Images.Count > 0 Then
        dcmCacheImg.CurrentIndex = 1
    End If
End Sub


Private Function GetRootHwnd() As Long
    GetRootHwnd = GetAncestor(hwnd, GA_ROOT)
End Function

Private Sub cboCacheMark_Change()
On Error GoTo errhandle
    Call LoadCacheImages(dtpCacheDate.value, cboCacheMark.Text)
Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "提示"
End Sub

Private Sub cboCacheMark_Click()
On Error GoTo errhandle
    Call LoadCacheImages(dtpCacheDate.value, cboCacheMark.Text)
Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "提示"
End Sub

Private Sub dcmCacheImg_DblClick()
On Error GoTo errhandle
    RaiseEvent OnDblClick
Exit Sub
errhandle:
    Call MsgboxH(GetRootHwnd, err.Description, vbOKOnly, "提示")
End Sub

Private Sub dcmCacheImg_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo errhandle
    Dim lngSelectIndex As Long
    Dim i As Long
    
    If dcmCacheImg.Images.Count <= 0 Then Exit Sub
    
    Select Case KeyCode
        Case 37     '左光标键盘
            lngSelectIndex = mlngSelImgIndex - 1
            If lngSelectIndex <= 0 Then Exit Sub
        Case 38    '上光标键
            lngSelectIndex = mlngSelImgIndex - dcmCacheImg.MultiColumns
            If lngSelectIndex <= 0 Then Exit Sub
        Case 39      '右光标键
            lngSelectIndex = mlngSelImgIndex + 1
            If lngSelectIndex > dcmCacheImg.Images.Count Then Exit Sub
        Case 40      '下光标键
            lngSelectIndex = mlngSelImgIndex + dcmCacheImg.MultiColumns
            If lngSelectIndex > dcmCacheImg.Images.Count Then Exit Sub
        Case 65
            If Shift = 2 Then
                Call SelectedAll  '按下全选
                lngSelectIndex = 0
                Exit Sub
            End If
            
        Case Else
            Exit Sub
    End Select
    
    For i = 1 To dcmCacheImg.Images.Count
        Call DrawBorder(dcmCacheImg.Images(i), 0)
    Next
        
    If lngSelectIndex > 0 Then
        Call DrawBorder(dcmCacheImg.Images(lngSelectIndex), ColorConstants.vbRed, True)
    End If
    
    mlngSelImgIndex = lngSelectIndex
     
Exit Sub
errhandle:
    Call MsgboxH(GetRootHwnd, err.Description, vbOKOnly, "提示")
End Sub

Private Sub dcmCacheImg_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errhandle
    RaiseEvent OnMouseDown(Button, Shift, X, Y)
Exit Sub
errhandle:
End Sub

Private Sub dcmCacheImg_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errhandle
    RaiseEvent OnMouseMove(Button, Shift, X, Y)
Exit Sub
errhandle:
End Sub

'Private Sub cmdRefresh_Click()
'On Error GoTo errhandle
'    Call LoadCacheMark(dtpCacheDate.value)
'Exit Sub
'errhandle:
'    MsgBoxH hWnd, err.Description, vbOKOnly, "提示"
'End Sub

Private Sub dcmCacheImg_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errhandle
    Dim i As Long
    
    If Button = 2 Then
        '鼠标右键
    Else
        mlngSelImgIndex = dcmCacheImg.ImageIndex(X, Y)
        
        If mlngSelImgIndex <= 0 Or mlngSelImgIndex > dcmCacheImg.Images.Count Then Exit Sub
        
        If Shift <> 2 Then
            For i = 1 To dcmCacheImg.Images.Count
                Call DrawBorder(dcmCacheImg.Images(i), 0)
            Next
        End If
            
        Call DrawBorder(dcmCacheImg.Images(mlngSelImgIndex), ColorConstants.vbRed, True)
    End If
    
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "提示"
End Sub

Private Sub dtpCacheDate_Change()
On Error GoTo errhandle
    Call LoadCacheMark(dtpCacheDate.value)
Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "提示"
End Sub

Private Sub dtpCacheDate_Click()
On Error GoTo errhandle
    Call LoadCacheMark(dtpCacheDate.value)
Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "提示"
End Sub

Private Sub UserControl_Initialize()
    dtpCacheDate.value = Now
End Sub

Private Sub UserControl_Resize()
    Dim lngRow As Integer
    Dim lngCol As Integer
    
On Error Resume Next
    Call InitFace
    
    Call ResizeRegion(GetPageRecordCount, _
                        dcmCacheImg.Width, dcmCacheImg.Height, lngRow, lngCol)
    
    dcmCacheImg.MultiRows = lngRow
    dcmCacheImg.MultiColumns = lngCol
End Sub


Public Sub Destory()

End Sub

Private Sub UserControl_Terminate()
    Call Destory
End Sub

Private Sub vscImages_Change()
On Error GoTo errhandle
    If dcmCacheImg.Images.Count <= 0 Then Exit Sub
    If vscImages.Max <= 0 Then Exit Sub
    dcmCacheImg.CurrentIndex = vscImages.value
Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "提示"
End Sub

Private Sub vscImages_Scroll()
On Error GoTo errhandle
    If dcmCacheImg.Images.Count <= 0 Then Exit Sub
    If vscImages.Max <= 0 Then Exit Sub
    dcmCacheImg.CurrentIndex = vscImages.value
Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "提示"
End Sub
