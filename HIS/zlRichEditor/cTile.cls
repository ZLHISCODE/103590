VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'#########################################################################
'##模 块 名：cTile.cls
'##创 建 人：吴庆伟
'##日    期：2005年4月8日
'##描    述：将一幅图片平铺到指定DC的区域内
'##要    求：
'#########################################################################
Option Explicit

Private m_lHdc As Long              '图片DC值
Private m_lHBmp As Long             '图片句柄（设备场景外）
Private m_lHBmpOld As Long          '旧图片句柄（内部）
Private m_pic As StdPicture         '存储图片内容
Private m_sFileName As String       '图片文件名
Private m_lXOriginOffset As Long    'X原点偏移量
Private m_lYOriginOffset As Long    'Y原点偏移量
Private m_lBitmapW As Long          '位图宽度
Private m_lBitmapH As Long          '位图高度

Private Const cTileErrorBase = 5600

 '图片DC值
 Public Property Get PicturehDC() As Long
    PicturehDC = m_lHdc
End Property

 'X原点偏移量
Public Property Get XOriginOffset() As Long
    XOriginOffset = m_lXOriginOffset
End Property

 'X原点偏移量
Public Property Let XOriginOffset(ByVal lPixels As Long)
    m_lXOriginOffset = lPixels
End Property

'Y原点偏移量
Public Property Get YOriginOffset() As Long
    YOriginOffset = m_lYOriginOffset
End Property

'Y原点偏移量
Public Property Let YOriginOffset(ByVal lPixels As Long)
    m_lYOriginOffset = lPixels
End Property

'位图宽度
Public Property Get BitmapWidth() As Long
    BitmapWidth = m_lBitmapW
End Property

'位图高度
Public Property Get BitmapHeight() As Long
    BitmapHeight = m_lBitmapH
End Property

'弹出错误消息函数，参数：错误号，消息体
Private Sub pErr(lNumber As Long, smsg As String)
    MsgBox "错误: " & smsg & "，" & lNumber, vbExclamation
End Sub

'图片文件名
Public Property Let FileName(ByVal sFIleName As String)
    '从一个文件调入图片
    If (m_sFileName <> sFIleName) Then
        '清除图片相关内容
        pClearUp
        If (pbLoadPicture(sFIleName)) Then
            m_sFileName = sFIleName
        End If
    End If
End Property

'图片文件名
Public Property Get FileName() As String
    FileName = m_sFileName
End Property

'获取图片内容
Public Property Get Picture() As StdPicture
    Set Picture = m_pic
End Property

'图片内容赋值
Public Property Let Picture(oPic As StdPicture)
   ' 从一个 StdPicture 对象调入图片
   pClearUp '清除位图
   If Not oPic Is Nothing Then
      If (pbEnsurePicture()) Then   '分配内存空间
          Set m_pic = oPic          '赋值
          If (Err.Number = 0) Then
              pbGetBitmapIntoDC     '同时分配相关句柄
          End If
      End If
   End If
End Property

Private Function pbEnsurePicture() As Boolean
'为图片分配内存空间，确保能够保存图片，确定返回 True，否则返回 False
On Error Resume Next
    pbEnsurePicture = True
    If (m_pic Is Nothing) Then
        Set m_pic = New StdPicture  '标准位图
        If (Err.Number <> 0) Then
            pErr 3, "无法为图片对象分配内存空间。"
            pbEnsurePicture = False
        Else
        End If
    End If
On Error GoTo 0
    Exit Function
End Function

Private Function pbLoadPictureFromFile(sFile As String) As Boolean
'调入指定文件名的图片，成功返回 True
On Error Resume Next
   Set m_pic = LoadPicture(sFile) '加载图片
   If (Err.Number <> 0) Then
       pErr 0, "载入图片失败: " & Err.Description
   Else
       pbLoadPictureFromFile = True
   End If
On Error GoTo 0
    Exit Function
End Function

Private Function pbLoadPicture(sFile As String) As Boolean
'函数作用：调入指定名称的图片，成功返回 True
    '先分配图片内存空间 pbEnsurePicture()
    If (pbEnsurePicture()) Then
        '正式载入图片
        If (pbLoadPictureFromFile(sFile)) Then
            '载入成功，则
            pbLoadPicture = pbGetBitmapIntoDC()
        End If
    End If
    
End Function

Private Function pbGetBitmapIntoDC() As Boolean
'函数作用： 获取位图信息到DC中，成功返回 True
Dim tB As BITMAP
Dim lHDC As Long, lHwnd As Long
Dim lHDCTemp As Long, lHBmpTempOld As Long

    ' 使得一个DC来容纳图片位图，这样我们才能通过它进行Blt操作。
    lHwnd = GetDesktopWindow()  '获取桌面（即屏幕，所有窗体、图标的绘制区域）的 Window 句柄。
    lHDC = GetDC(lHwnd)         '获取该句柄的 DC ，参数为一个 Window 句柄。
    m_lHdc = CreateCompatibleDC(lHDC)   '创建兼容的内存设备场景（即内存图像）。
    lHDCTemp = CreateCompatibleDC(lHDC) '
    If (m_lHdc <> 0) Then
        ' 获取位图尺寸
        ' GetObjectAPI 取得对指定对象进行说明的一个结构
        GetObjectAPI m_pic.Handle, LenB(tB), tB
        m_lBitmapW = tB.bmWidth     '获取图片宽度
        m_lBitmapH = tB.bmHeight    '获取图片高度
        
        'SelectObject 选择指定对象到设备场景中，自动覆盖旧的同类对象。
        '参数1：DC句柄，参数2：对象的句柄。
        lHBmpTempOld = SelectObject(lHDCTemp, m_pic.Handle)
        
        '创建一个设备相关位图
        '参数1：设备场景句柄；参数2：图片宽度；参数3：图片高度
        m_lHBmp = CreateCompatibleBitmap(lHDC, m_lBitmapW, m_lBitmapH)
        m_lHBmpOld = SelectObject(m_lHdc, m_lHBmp)  '选取该位图
        
        '赋值矩形图片块到目标上下文中
        BitBlt m_lHdc, 0, 0, m_lBitmapW, m_lBitmapH, lHDCTemp, 0, 0, vbSrcCopy
        
        '选择图片
        SelectObject lHDCTemp, lHBmpTempOld
        '删除设备场景
        DeleteDC lHDCTemp
        
        If (m_lHBmpOld <> 0) Then
            pbGetBitmapIntoDC = True
            If (m_sFileName = "") Then
               m_sFileName = "PICTURE"
            End If
        Else
            pClearUp    '清除图片
            pErr 2, "无法选取图片到设备场景中。"
        End If
    Else
        pErr 1, "无法创建兼容的设备场景。"
    End If
    '释放设备场景
    ReleaseDC lHwnd, lHDC
    
End Function

'调色板（返回当前图片内容）
Public Property Get Palette() As StdPicture
    Set Palette = m_pic
End Property

Private Sub pClearUp()
'函数作用：清除图片相关资源（几个句柄:m_lHdc/m_lHBmpOld/m_lHBmp）

    ' 清除文件名的引用
    m_sFileName = ""
    ' 如果我们已经有了一个DC，那么清除它
    If (m_lHdc <> 0) Then
        ' 选取一个DC外位图
        If (m_lHBmpOld <> 0) Then
            SelectObject m_lHdc, m_lHBmpOld
            ' 原始位图无需删除，因为它由 m_pic 所有。
        End If
        If (m_lHBmp <> 0) Then
            DeleteObject m_lHBmp    '删除指定句柄的图形对象
        End If
        ' 移除DC
        DeleteDC m_lHdc
    End If
End Sub

' 平铺图片到指定区域
Public Sub TileArea( _
        ByRef hdc As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal Width As Long, _
        ByVal Height As Long _
    )
Dim lSrcX As Long
Dim lSrcY As Long
Dim lSrcStartX As Long
Dim lSrcStartY As Long
Dim lSrcStartWidth As Long
Dim lSrcStartHeight As Long
Dim lDstX As Long
Dim lDstY As Long
Dim lDstWidth As Long
Dim lDstHeight As Long

    lSrcStartX = ((X + m_lXOriginOffset) Mod m_lBitmapW)
    lSrcStartY = ((Y + m_lYOriginOffset) Mod m_lBitmapH)
    lSrcStartWidth = (m_lBitmapW - lSrcStartX)
    lSrcStartHeight = (m_lBitmapH - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    
    lDstY = Y
    lDstHeight = lSrcStartHeight
    
    Do While lDstY < (Y + Height)
        If (lDstY + lDstHeight) > (Y + Height) Then
            lDstHeight = Y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = X
        lSrcX = lSrcStartX
        Do While lDstX < (X + Width)
            If (lDstX + lDstWidth) > (X + Width) Then
                lDstWidth = X + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            '绘制矩形图片内容
            BitBlt hdc, lDstX, lDstY, lDstWidth, lDstHeight, m_lHdc, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = m_lBitmapW
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = m_lBitmapH
    Loop
End Sub


Private Sub Class_Terminate()
    '确定所有GDI对象都被释放
    pClearUp
    '清除图片
    Set m_pic = Nothing
End Sub



