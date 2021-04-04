VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmPACSImg 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6396
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   ScaleHeight     =   6396
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin DicomObjects.DicomViewer DViewer 
      Height          =   4875
      Left            =   105
      TabIndex        =   0
      Top             =   15
      Width           =   2535
      _Version        =   262147
      _ExtentX        =   4471
      _ExtentY        =   8599
      _StockProps     =   35
      BackColor       =   14737632
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   90
      Top             =   5535
      _Version        =   589884
      _ExtentX        =   508
      _ExtentY        =   508
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmPACSImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngAdviceID As Long      '当前编辑的医嘱id
Private mlngCurIndex As Long      '当前选择的图像序号
Private WithEvents mfrmShow As frmPacsImgShow
Attribute mfrmShow.VB_VarHelpID = -1
Private mlngModule As Long

'公共事件
Public Event RequestRightMenu(ByRef cbsThis As Object)
Public Event InsertPicture(ByRef pic As StdPicture, ByVal strUid As String, ByVal lngAdviceID As Long)

Private Sub ConfigImgDisplayFormat(ByVal lngPageRecord As Long)
'配置图像显示格式
    Dim iRows As Integer
    Dim iCols As Integer
    
    ResizeRegion lngPageRecord, DViewer.Width, DViewer.Height, iRows, iCols

    DViewer.MultiColumns = iCols
    DViewer.MultiRows = iRows
End Sub

Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'功能：创建本地目录
'参数： strDir－－本地目录
'返回：无
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next

    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir

    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Function GetCacheDir() As String
'获取缓存目录
    GetCacheDir = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
End Function

Private Function LoadAllCaptureImage(ByVal lngAdviceID As Long, dcmViewer As DicomViewer) As Boolean
'加载预览图像到界面
    Dim strTmpFile As String
    Dim strCachePath As String
    
    Dim curImage As DicomImage
    
    Dim objFile As New Scripting.FileSystemObject
    
    Dim Inet1 As New cFTP
    Dim Inet2 As New cFTP
    
    Dim strImgInstanceUid As String
    Dim strCurInstanceUids As String
    Dim blnIsAddImage As Boolean
    Dim objImgInfo As Object
    
    Dim strSQL As String
    Dim rsCurImageData As ADODB.Recordset

    strSQL = "Select rownum as 顺序号,A.图像UID,c.姓名,c.性别,c.年龄, A.图像号,D.FTP用户名 As User1,D.FTP密码 As Pwd1," & _
            "D.IP地址 As Host1,'/'||D.Ftp目录||'/' As Root1,D.共享目录 as 共享目录1,D.共享目录用户名 as 共享目录用户名1,D.共享目录密码 as 共享目录密码1," & _
            "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/') " & _
            "||C.检查UID||'/'||A.图像UID As URL,d.设备号 as 设备号1," & _
            "E.FTP用户名 As User2,E.FTP密码 As Pwd2," & _
            "E.IP地址 As Host2,'/'||E.Ftp目录||'/' As Root2,E.共享目录 as 共享目录2,E.共享目录用户名 as 共享目录用户名2,E.共享目录密码 as 共享目录密码2," & _
            "E.设备号 as 设备号2, A.图像UID,C.检查UID,B.序列UID,A.动态图,A.编码名称,A.采集时间, A.录制长度 " & _
            "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
            "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) And nvl(A.动态图,0) = 0 and c.医嘱ID = [1]"
    Set rsCurImageData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID)
    
    blnIsAddImage = False
    
    LoadAllCaptureImage = False
    
    If rsCurImageData.RecordCount <= 0 Then
        dcmViewer.Images.Clear
        
        dcmViewer.MultiColumns = 1
        dcmViewer.MultiRows = 1
        
        Exit Function
    End If
    
    strCurInstanceUids = ""
        
    '配置图像显示格式
    If rsCurImageData.RecordCount <> dcmViewer.Images.Count Then
        Call ConfigImgDisplayFormat(rsCurImageData.RecordCount)
    End If
        
    '创建本地图像缓存目录
    strCachePath = GetCacheDir
    MkLocalDir strCachePath & objFile.GetParentFolderName(NVL(rsCurImageData("URL")))
    
    Do While Not rsCurImageData.EOF
        '循环加载图像到DicomViewer中
        strImgInstanceUid = NVL(rsCurImageData!图像UID)
        
        If InStr(strCurInstanceUids, strImgInstanceUid) <= 0 Then
            
            blnIsAddImage = True
            
            '设置音视频的显示文件，如果为音视频文件时，该过程将不从服务器中直接下载数据文件
            If NVL(rsCurImageData!动态图, 0) = 0 Then
                strTmpFile = strCachePath & NVL(rsCurImageData("URL"))
            End If
            
            If Dir(strTmpFile) = vbNullString Then
                '本地缓存图像不存在，则读取FTP图像
                '建立FTP连接
                If NVL(rsCurImageData("设备号1")) <> vbNullString And Inet1.hConnection = 0 Then
                    If Inet1.FuncFtpConnect(NVL(rsCurImageData("Host1")), NVL(rsCurImageData("User1")), NVL(rsCurImageData("Pwd1"))) = 0 Then
                        If NVL(rsCurImageData("设备号2")) <> vbNullString Then
                            If Inet2.FuncFtpConnect(NVL(rsCurImageData("Host2")), NVL(rsCurImageData("User2")), NVL(rsCurImageData("Pwd2"))) = 0 Then
                                Exit Function
                            End If
                        Else
                            Exit Function
                        End If
                    End If
                End If
                
                If Inet1.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsCurImageData("Root1")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL"))) <> 0 Then
                    '从设备号1提取图像失败，则从设备号2提取图像
                    If NVL(rsCurImageData("设备号2")) <> vbNullString Then
                        If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect NVL(rsCurImageData("Host2")), NVL(rsCurImageData("User2")), NVL(rsCurImageData("Pwd2"))
                        Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsCurImageData("Root2")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL")))
                    End If
                End If
            End If
    
            If Dir(strTmpFile) <> vbNullString Then
                If NVL(rsCurImageData!动态图, 0) = 0 Then
                    Err.Clear
                    On Error Resume Next
                    Set curImage = dcmViewer.Images.ReadFile(strTmpFile)
                    
                    If Err.Number <> 0 Then
                        Set curImage = dcmViewer.Images.AddNew
                        Call curImage.FileImport(strTmpFile, "JPG")
                    End If
                    
                    curImage.Tag = NVL(rsCurImageData("图像UID")) & ".jpg"
                    
                    With curImage
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                End If
                
                '取消自动剪影,因为DicomObjects控件本身对处理剪影有BUG，存在（0028，6100）时，会自动对图像进行剪影，
                '导致晋煤的DSA图像不能正常显示
                '虽然设置图像的mask=0 ,可以取消剪影，但是每次图像被添加到新的Dicomimages之后，自动又将mask设置成1了，
                '这样在程序中无法很好的控制，因此直接去掉（0028，6100）这个属性。
                If Not IsNull(curImage.Attributes(&H28, &H6100).Value) Then
                    curImage.Attributes.Remove &H28, &H6100
                End If
            End If
        End If
        
        rsCurImageData.MoveNext
    Loop
    
    UpdateSelectIndex 1
    
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect
    
    LoadAllCaptureImage = True
End Function

Private Sub UpdateSelectIndex(ByVal lngSelectIndex As Long)
'配置图像的选中索引
    Dim blnIsValidIndex As Boolean
    
    blnIsValidIndex = IIf(lngSelectIndex > 0 And lngSelectIndex <= Me.DViewer.Images.Count, True, False)
    
    If Not blnIsValidIndex Then Exit Sub

    If blnIsValidIndex Then DViewer.Images(lngSelectIndex).BorderColour = vbRed
    If mlngCurIndex = lngSelectIndex Then Exit Sub

    If mlngCurIndex > 0 And mlngCurIndex <= DViewer.Images.Count Then
        DViewer.Images(mlngCurIndex).BorderColour = vbWhite
    End If

    mlngCurIndex = lngSelectIndex
End Sub

Private Function LoadSelectReportImage(ByVal lngAdviceID As Long, dcmViewer As DicomViewer) As Boolean
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim aryFiles() As String    '报告图像数组
    Dim strFiles As String      '按分号分隔的成功下载的文件
    
    Dim cFtpNet As New cFTP
    Dim strVirtualPath As String
    Dim strLocalPath As String
    Dim intCount As Integer
     
    LoadSelectReportImage = False
    
    '先根据报告图象字段获取报告图信息，没有值则读取所有检查图像
    strSQL = "Select To_Char(L.接收日期, 'yyyymmdd') As 子目录, L.检查uid, L.报告图象, A1.Ftp目录 As Root1, A1.Ip地址 As Ip1," & vbNewLine & _
            "       A1.FTP用户名 As Usr1, A1.FTP密码 As Pwd1, A2.Ftp目录 As Root2, A2.Ip地址 As Ip2, A2.FTP用户名 As Usr2, A2.FTP密码 As Pwd2" & vbNewLine & _
            "From 影像检查记录 L, 影像设备目录 A1, 影像设备目录 A2" & vbNewLine & _
            "Where L.位置一 = A1.设备号(+) And L.位置二 = A2.设备号(+) And L.报告图象 Is Not Null And L.医嘱id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取信息", lngAdviceID)

    If rsTemp.RecordCount <= 0 Then
        dcmViewer.Images.Clear
        
        dcmViewer.MultiColumns = 1
        dcmViewer.MultiRows = 1
        
        Exit Function
    End If
    
    aryFiles = Split("" & rsTemp!报告图象, ";")
    If UBound(aryFiles) < 0 Then
        dcmViewer.Images.Clear
        
        dcmViewer.MultiColumns = 1
        dcmViewer.MultiRows = 1
        
        Exit Function
    End If
        
    '创建本地存储目录
    Err = 0: On Error Resume Next
    strLocalPath = App.Path & "\TmpImage\" & rsTemp!子目录
    If objFileSystem.FolderExists(strLocalPath) = False Then objFileSystem.CreateFolder strLocalPath
    If objFileSystem.FolderExists(strLocalPath) = False Then Exit Function
    
    strLocalPath = strLocalPath & "\" & rsTemp!检查uid
    If objFileSystem.FolderExists(strLocalPath) = False Then objFileSystem.CreateFolder strLocalPath
    If objFileSystem.FolderExists(strLocalPath) = False Then Exit Function
        
    '判断连接的有效性，并从连接下载文件
    strFiles = ""
    If "" & rsTemp!Ip1 <> "" Then
        If cFtpNet.FuncFtpConnect("" & rsTemp!Ip1, "" & rsTemp!Usr1, "" & rsTemp!pwd1) <> 0 Then
            strVirtualPath = rsTemp!Root1 & "/" & rsTemp!子目录 & "/" & rsTemp!检查uid
            For intCount = 0 To UBound(aryFiles)
                If Dir(strLocalPath & "\" & Trim(aryFiles(intCount))) <> "" Then
                    strFiles = strFiles & ";" & strLocalPath & "\" & Trim(aryFiles(intCount))
                    aryFiles(intCount) = ""
                Else
                    If cFtpNet.FuncDownloadFile(strVirtualPath, strLocalPath & "\" & Trim(aryFiles(intCount)), Trim(aryFiles(intCount))) = 0 Then
                        If Dir(strLocalPath & "\" & Trim(aryFiles(intCount))) <> "" Then
                            strFiles = strFiles & ";" & strLocalPath & "\" & Trim(aryFiles(intCount))
                            aryFiles(intCount) = ""
                        End If
                    End If
                End If
            Next
        End If
        cFtpNet.FuncFtpDisConnect
    End If
    
    If strFiles <> "" Then strFiles = Mid(strFiles, 2)
    If UBound(Split(strFiles, ";")) <> UBound(aryFiles) And "" & rsTemp!Ip2 <> "" Then
        If cFtpNet.FuncFtpConnect("" & rsTemp!Ip2, "" & rsTemp!Usr2, "" & rsTemp!pwd2) <> 0 Then
            strVirtualPath = rsTemp!Root2 & "/" & rsTemp!子目录 & "/" & rsTemp!检查uid
            For intCount = 0 To UBound(aryFiles)
                If aryFiles(intCount) <> "" Then
                    If cFtpNet.FuncDownloadFile(strVirtualPath, strLocalPath & "\" & Trim(aryFiles(intCount)), Trim(aryFiles(intCount))) = 0 Then
                        If Dir(strLocalPath & "\" & Trim(aryFiles(intCount))) <> "" Then
                            strFiles = strFiles & ";" & strLocalPath & "\" & Trim(aryFiles(intCount))
                        End If
                    End If
                End If
            Next
        End If
        cFtpNet.FuncFtpDisConnect
    End If
    
    If strFiles <> "" Then
        If Left(strFiles, 1) = ";" Then strFiles = Mid(strFiles, 2)
    End If
    
    '将获得的文件装入
    Dim curImage As DicomImage, iRows As Integer, iCols As Integer, strCurName As String
    aryFiles = Split(strFiles, ";")
    With Me.DViewer
        If .Images.Count > 0 Then strCurName = .Images(mlngCurIndex - 1).Tag '记下当前所选
        mlngCurIndex = 0
        .Images.Clear
        For intCount = 0 To UBound(aryFiles)
            Set curImage = New DicomImage
            curImage.FileImport aryFiles(intCount), "JPG"
            curImage.BorderStyle = 6: curImage.BorderWidth = 1: curImage.BorderColour = vbWhite
            .Images.Add curImage
            .Images(intCount + 1).Tag = gobjFSO.GetFileName(aryFiles(intCount))                    '记下文件名作标记
            If strCurName = aryFiles(intCount) Then mlngCurIndex = intCount + 1
        Next
        
        If .Images.Count > 0 Then
            If mlngCurIndex = 0 Then
                mlngCurIndex = 1
                If strCurName <> "" Then Unload mfrmShow                '刷新之前有的图,刷新之后没有,表明可能被删除
            End If
            .CurrentIndex = 1
            .Images(mlngCurIndex).BorderColour = vbRed
            
            ResizeRegion .Images.Count, .Width, .Height, iRows, iCols '调整排列
            .MultiColumns = iCols: .MultiRows = iRows
        Else
            Unload mfrmShow
        End If
    End With
    
End Function

Public Function zlRefresh(ByVal lngAdviceID As Long, ByVal lngModule As Long) As Boolean
    '获取报告图像到本地，并刷新显示
On Error GoTo errHand
        
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim aryFiles() As String    '报告图像数组
    Dim strFiles As String      '按分号分隔的成功下载的文件
    
    Dim cFtpNet As New cFTP
    Dim strVirtualPath As String
    Dim strLocalPath As String
    Dim intCount As Integer
     
    mlngAdviceID = lngAdviceID
    mlngModule = lngModule
    
    '创建本地根目录
    If objFileSystem.FolderExists(App.Path & "\TmpImage\") = False Then objFileSystem.CreateFolder App.Path & "\TmpImage\"
    If objFileSystem.FolderExists(App.Path & "\TmpImage\") = False Then zlRefresh = False: Exit Function
     
    '根据不同模块，使用不同的报告图读取方式
    If mlngModule = 1290 Then
        zlRefresh = LoadSelectReportImage(lngAdviceID, DViewer)
    Else
        zlRefresh = LoadAllCaptureImage(lngAdviceID, DViewer)
    End If

    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'---------------------------------------------------
'以下是窗体空间事件处理
'---------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Insert
        Call DViewer_DblClick
    Case conMenu_View_Refresh
        Call Me.zlRefresh(mlngAdviceID, mlngModule)
    Case conMenu_View_Option
        If mfrmShow.Visible Then
            Unload mfrmShow
        Else
            mfrmShow.Show , Me
            Call DViewer_Click
        End If
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Left = -120
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Dim iRows As Integer, iCols As Integer
    Call Me.cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    Err = 0: On Error Resume Next
    With Me.DViewer
        iRows = .MultiRows: iCols = .MultiColumns
        .Left = lngScaleLeft + 120: .Top = lngScaleTop
        .Width = lngScaleRight - .Left: .Height = lngScaleBottom - .Top
        If .Images.Count > 0 Then
            ResizeRegion .Images.Count, .Width, .Height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows
        End If
    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Insert
        Control.Enabled = (mlngCurIndex > 0)
    Case conMenu_View_Option
        Control.Enabled = (Me.DViewer.Images.Count > 0)
        If (Control.Enabled = False Or Me.Visible = False) And mfrmShow.Visible Then Unload mfrmShow
        Control.Checked = mfrmShow.Visible
    End Select
End Sub

Private Sub DViewer_Click()
    Dim pic As StdPicture, picUid As String
    If mfrmShow.Visible = False Then Exit Sub
    If mlngCurIndex > 0 Then
        Set pic = Me.DViewer.Images(mlngCurIndex).Picture
        picUid = Me.DViewer.Images(mlngCurIndex).Tag
    End If
    
    If Not (pic Is Nothing) Then
        Set mfrmShow.imgShow.Picture = pic
        mfrmShow.imgShow.Tag = picUid
    Else
        Set mfrmShow.imgShow.Picture = Nothing
        mfrmShow.imgShow.Tag = ""
    End If
End Sub

Private Sub DViewer_DblClick()
    Dim pic As StdPicture, picUid As String
    If mlngCurIndex > 0 Then
        Set pic = Me.DViewer.Images(mlngCurIndex).Picture
        picUid = Me.DViewer.Images(mlngCurIndex).Tag
    End If
    If Not pic Is Nothing Then RaiseEvent InsertPicture(pic, picUid, mlngAdviceID)
End Sub

Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim i As Integer
    If Button <> 1 Then Exit Sub
    With DViewer
        i = .ImageIndex(X, Y)
        If i > 0 And i <= .Images.Count And i <> mlngCurIndex Then
            .Images(mlngCurIndex).BorderColour = vbWhite
            .Images(i).BorderColour = vbRed
            mlngCurIndex = i
        End If
    End With
End Sub

Private Sub Form_Load()
    Set mfrmShow = New frmPacsImgShow
    
    Dim cbrControl As CommandBarControl
    '-----------------------------------------------------
    '内部菜单工具栏定义
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.Position = xtpBarBottom
    Me.cbsThis.ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    With Me.cbsThis.ActiveMenuBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Insert, "添入报告(&S)"): cbrControl.STYLE = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.flags = xtpFlagRightAlign: cbrControl.STYLE = xtpButtonCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Option, "大图(&B)"): cbrControl.flags = xtpFlagRightAlign: cbrControl.STYLE = xtpButtonCaption
        cbrControl.Checked = False
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmShow Is Nothing Then Unload mfrmShow
    Set mfrmShow = Nothing
End Sub
 




Private Sub mfrmShow_DblClick(pic As stdole.StdPicture, ByVal strUid As String)
    RaiseEvent InsertPicture(pic, strUid, mlngAdviceID)
End Sub
