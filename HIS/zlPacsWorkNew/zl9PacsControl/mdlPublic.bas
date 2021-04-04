Attribute VB_Name = "mdlPublic"
Option Explicit

Public Const CON_STR_HINT_TITLE As String = "提示"

'图像标注
Public Const m_LabelTag_Circle = "NumberCircle"
Public Const m_LabelTag_Back = "NumberBak"
Public Const m_LabelTag_Number = "Number"


Public gcnOracle As ADODB.Connection
Public gobjOwner As Object
Public glngOwnerHwnd As Long
Public gblnOpenDebug As Boolean
Public glngSys As Long
Public glngMoudle As Long

Private gstrDebugPath As String

Public Function GetAppPath() As String
    If gstrDebugPath = "" Then
        If App.LogMode = 0 Then
            gstrDebugPath = "C:\Appsoft\Apply"
        Else
            gstrDebugPath = Replace(App.Path & "\", "\\", "")
        End If
    End If
    
    GetAppPath = gstrDebugPath
End Function

Public Function DynamicCreate(ByVal strclass As String, ByVal strCaption As String) As Object
'动态创建对象
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
   
    If err <> 0 Then
        MsgBox strCaption & "组件创建失败，请联系管理员检查是否正确安装!", vbInformation, CON_STR_HINT_TITLE
        Set DynamicCreate = Nothing
    End If
    err.Clear
End Function


Public Sub BUGEX(ByVal strDebug As String, Optional ByVal blnIsForce As Boolean = False)
    If gblnOpenDebug Or blnIsForce Then
        OutputDebugString Format(Now, "mmddhhmmss") & " |-> " & strDebug
    End If
End Sub

Public Function GetCacheDir() As String
'获取缓存目录
    GetCacheDir = GetAppPath & "\TmpImage\"
End Function
'
'
Public Function GetResourceDir() As String
'获取资源目录
    GetResourceDir = GetAppPath & "\..\附加文件\"
End Function


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

Public Sub AddVideoLabelToDicomImage(dcmImage As DicomImage, ByVal strCaptureTimeText As String, _
    ByVal strTimeLenText As String, ByVal strEncoderName As String)
    '功能:添加label
    '参数:dcmImage：dicom图像
    '     strCaption： label文本
    Dim labCaption As New DicomLabel
    
    labCaption.LabelType = doLabelText
    '不显示编码器的名称
    labCaption.Text = strCaptureTimeText & vbCrLf & strTimeLenText '& vbCrLf & strEncoderName
    labCaption.Font.Bold = True
    labCaption.Font.Name = "宋体"
    labCaption.Font.Size = 10
    labCaption.ForeColour = vbYellow
    labCaption.AutoSize = False

    
    labCaption.Left = 0
    labCaption.Top = 0
    
    Call dcmImage.Labels.Add(labCaption)
End Sub


Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer)
'-----------------------------------------------------------------------------
'功能：根据输入的图像数量，图像区域的宽度和高度，返回最佳的图像排列行数和列数
'参数： ImageCount－－图像数量
'       RegionWidth--图像显示区域的宽度
'       RegionHeight--图像显示区域的高度
'       Rows－－[返回]最佳行数
'       Cols－－[返回]最佳列数
'返回：返回最佳行数Rows，最佳列数Cols
'-----------------------------------------------------------------------------
    Dim iCols As Integer, iRows As Integer
    Dim iBase As Integer, blnDoLoop As Integer
    Dim lngFreeCount As Long
    
    If RegionHeight = 0 Then RegionHeight = 1
    If RegionWidth = 0 Then RegionWidth = 1
    
    On Error GoTo err
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))

    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    
    '当图像格式为如下等形式时，需要对行列进行修正
    
    '格式1：
    '图1  图2  图3  图4
    '图5  图6  图7  图8
    '空1  空2  空3  空4
    
    '格式2：
    '图1  图2  图3  图4
    '图5  图6  图7  图8
    '图9  空1  空2  空3
    
    lngFreeCount = iRows * iCols - ImageCount
    Do While lngFreeCount >= iCols Or lngFreeCount >= iRows
        If lngFreeCount >= iCols Then
            iRows = iRows - 1
        Else
            iCols = iCols - 1
        End If
        
        lngFreeCount = iRows * iCols - ImageCount
    Loop
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    
    Do While iRows * iCols < ImageCount
        If RegionWidth / iCols > RegionHeight > iRows Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    
    '再次修正行列数
    lngFreeCount = iRows * iCols - ImageCount
    Do While lngFreeCount >= iCols Or lngFreeCount >= iRows
        If lngFreeCount >= iCols Then
            iRows = iRows - 1
        Else
            iCols = iCols - 1
        End If
        
        lngFreeCount = iRows * iCols - ImageCount
    Loop
    
    Rows = iRows: Cols = iCols
err:
End Sub

Public Function GetDecryptionPassW(ByVal strPassWord As String) As String
'如果密码已经加密，则需解密加密密码
    Dim strDecryptionPassW As String
    Dim objFrp As New clsFtp
    
    GetDecryptionPassW = strPassWord
    
    If Len(strPassWord) >= 3 Then
        If Mid(strPassWord, 1, 1) & Mid(strPassWord, 3, 1) & Mid(strPassWord, Len(strPassWord), 1) = "★※★" Then
            strDecryptionPassW = Mid(strPassWord, 2)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, Len(strDecryptionPassW) - 1)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, 1) & Mid(strDecryptionPassW, 3)
            strDecryptionPassW = objFrp.GetDecryptionPassW(strDecryptionPassW)
            
            GetDecryptionPassW = strDecryptionPassW
        End If
    End If
End Function

Public Function GetNewLabel(lType As Integer, lLeft As Integer, lTop As Integer, lWidth As Integer, lHeight As Integer) As DicomLabel
'------------------------------------------------
'功能：生成一个LABEL对象，并对其做初始化。
'参数：lType--标注的类型；lLeft--标注的Left值；lTop--标注的Top值；lWidth--标注的Width值；lHeight--标注的Height值。
'返回：新生成的标注。
'------------------------------------------------
    Dim l As New DicomLabel
    l.LabelType = lType
    l.ImageTied = True
    l.Left = lLeft
    l.Top = lTop
    l.Width = lWidth
    l.Height = lHeight
    l.Margin = 0
    l.AutoSize = True
    l.FontSize = 10
    l.LineWidth = 1
    'l.ForeColour = vbBlack
    l.XOR = True
    
    Set GetNewLabel = l
End Function

Public Function GetUserInfo() As String
'功能：获取登陆用户信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    If Not rsTmp.EOF Then
        GetUserInfo = IIf(IsNull(rsTmp!用户名), "", rsTmp!用户名)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
