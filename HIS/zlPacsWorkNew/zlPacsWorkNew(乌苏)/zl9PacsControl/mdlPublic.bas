Attribute VB_Name = "mdlPublic"
Option Explicit


Public Const CON_STR_HINT_TITLE As String = "提示"

Public gcnOracle As ADODB.Connection
Public gobjOwner As Object
Public glngOwnerHwnd As Long
Public gblnOpenDebug As Boolean

Public gobjGetImage() As Object       'zlPacsGetImage.clsPacsGetIamge
Public gblnUseActivexLoad As Boolean

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

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function


Public Sub BUGEX(ByVal strDebug As String, Optional ByVal blnIsForce As Boolean = False)
    If gblnOpenDebug Or blnIsForce Then
        OutputDebugString Format(Now, "mmddhhmmss") & " |-> " & strDebug
    End If
End Sub

Public Function GetCacheDir() As String
'获取缓存目录
    GetCacheDir = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
End Function
'
'
Public Function GetResourceDir() As String
'获取资源目录
    GetResourceDir = IIf(Len(App.Path) > 3, App.Path & "\..\附加文件\", App.Path & "..\附加文件\")
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


Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'功能：显示或隐藏一个窗体的标题栏
'参数：blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(objForm.hWnd, vRect)
    lngStyle = GetWindowLong(objForm.hWnd, GWL_STYLE)
    If blnCaption Then
        lngStyle = lngStyle Or WS_CAPTION Or WS_THICKFRAME
        If objForm.ControlBox Then lngStyle = lngStyle Or WS_SYSMENU
        If objForm.MaxButton Then lngStyle = lngStyle Or WS_MAXIMIZEBOX
        If objForm.MinButton Then lngStyle = lngStyle Or WS_MINIMIZEBOX
    Else
        If blnBorder Then
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
        Else
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
        End If
    End If
    SetWindowLong objForm.hWnd, GWL_STYLE, lngStyle
    SetWindowPos objForm.hWnd, 0, vRect.Left, vRect.Top, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
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
    
'    If ImageCount <> 0 Then
'        If Rows * Cols > ImageCount Then
'            iBase = 6
'            blnDoLoop = True
'
'            While blnDoLoop
'                iBase = iBase - 1
'
'                If ImageCount Mod iBase = 0 Then
'                    blnDoLoop = False
'                End If
'            Wend
'
'
'            If RegionWidth > RegionHeight Then
'                If ImageCount / iBase > iBase Then
'                    Cols = ImageCount / iBase
'                    Rows = iBase
'                Else
'                    Rows = ImageCount / iBase
'                    Cols = iBase
'                End If
'            Else
'                If ImageCount / iBase > iBase Then
'                    Cols = iBase
'                    Rows = ImageCount / iBase
'                Else
'                    Rows = iBase
'                    Cols = ImageCount / iBase
'                End If
'            End If
'        End If
'    End If
err:
End Sub

