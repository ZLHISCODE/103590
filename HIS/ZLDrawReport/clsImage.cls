VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################


Public Enum IMAGESTYLE
    自动缩放
    自动拉伸
End Enum
Private mblnClear As Boolean
Private mobjDraw As Object
Private mobjDrawDest As Object
Private mbytImageStyle As IMAGESTYLE
Private mlngGraphics As Long
Private mlngGraphicsDest As Long
Private mlngImgHandle As Long
Private mstrFile As String
Private mlngImageWidth As Long
Private mlngImageHeight As Long

'######################################################################################################################

Public Function InitGDI() As Boolean

    InitGDI = True
    
End Function

Public Function InitImage(ByVal objDraw As Object) As Boolean
    Set mobjDraw = objDraw
    
    Call GdipCreateFromHDC(mobjDraw.hDC, mlngGraphics)
    
    InitImage = (mlngGraphics > 0)
    
End Function


Public Sub ShowImage(ByVal strFile As String, Optional ByVal vKind As IMAGESTYLE = 自动拉伸, Optional ByVal blnClear As Boolean = True, _
                    Optional ByVal sglX As Single = 0, Optional ByVal sglY As Single = 0, Optional ByVal sglW As Single = 0, Optional ByVal sglH As Single = 0, _
                    Optional ByVal rfType As RotateFlipType = RotateNoneFlipNone)

    mbytImageStyle = vKind
    mblnClear = blnClear

    If strFile = "" Then Exit Sub
    If Dir(strFile) = "" Then Exit Sub

    If sglW = 0 Then sglW = mobjDraw.Width
    If sglH = 0 Then sglH = mobjDraw.Height

    Call ShowNormalImage(strFile, sglX / 15, sglY / 15, sglW / 15, sglH / 15, rfType)
    
End Sub

Public Sub DrawImage(Optional lngGap As Long, Optional bytCaptionStyle As Byte)
    'lngGap 当有文字的时候，文字所占高度
    'bytCaptionStyle 当有文字的时候，文字的对齐方式  1：底端对齐  2：顶端对其
    Call DrawNormalImage(lngGap, bytCaptionStyle)
End Sub

Public Function LoadImageFile(ByVal strFile As String) As Boolean
    
    mstrFile = strFile
    Call GdipLoadImageFromFile(StrPtr(strFile), mlngImgHandle)

    Call GdipGetImageWidth(mlngImgHandle, mlngImageWidth)    '
    Call GdipGetImageHeight(mlngImgHandle, mlngImageHeight)    '

    LoadImageFile = (mlngImgHandle > 0)
        
End Function


Public Function DisposeImage() As Boolean
    
    If mlngImgHandle > 0 Then
        Call GdipDisposeImage(mlngImgHandle)
        mlngImgHandle = 0
    End If

    If mlngGraphics > 0 Then
        Call GdipDeleteGraphics(mlngGraphics)
        mlngGraphics = 0
    End If

    Set mobjDraw = Nothing
    
    DisposeImage = True
    
End Function

Public Property Get ImageWidth() As Long
    ImageWidth = mlngImageWidth
End Property

Public Property Get ImageHeight() As Long
    ImageHeight = mlngImageHeight
End Property

'######################################################################################################################
Private Sub DrawNormalImage(Optional lngGap As Long, Optional bytCaptionStyle As Byte)

    Dim W As Long
    Dim H As Long
    Dim lngW As Long
    Dim lngH As Long
    Dim sngW As Single
    Dim sngH As Single
    
    If mobjDraw Is Nothing Then Exit Sub
    
    Select Case mbytImageStyle
    '------------------------------------------------------------------------------------------------------------------
    Case IMAGESTYLE.自动缩放

        H = ImageHeight
        W = ImageWidth

        If W <= mobjDraw.ScaleWidth / 15 And H <= mobjDraw.ScaleHeight / 15 Then
            lngW = W: lngH = H
        Else
            sngW = W / (mobjDraw.ScaleWidth / 15)
            sngH = H / (mobjDraw.ScaleHeight / 15)
            If sngW > sngH Then
                lngW = W / sngW: lngH = H / sngW
            Else
                lngW = W / sngH: lngH = H / sngH
            End If
        End If
        If lngGap > 0 And bytCaptionStyle > 0 Then
            Select Case bytCaptionStyle
            Case 1 '文字底端对齐,即图片在上面
                Call GdipDrawImageRect(mlngGraphics, mlngImgHandle, (mobjDraw.Width / 15 - lngW) / 2, (mobjDraw.Height / 15 - lngH - lngGap) / 2, lngW, lngH)
            Case 2
                Call GdipDrawImageRect(mlngGraphics, mlngImgHandle, (mobjDraw.Width / 15 - lngW) / 2, (mobjDraw.Height / 15 - lngH - lngGap) / 2 + lngGap, lngW, lngH)
            End Select
            
        Else
            Call GdipDrawImageRect(mlngGraphics, mlngImgHandle, (mobjDraw.Width / 15 - lngW) / 2, (mobjDraw.Height / 15 - lngH) / 2, lngW, lngH)
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case IMAGESTYLE.自动拉伸

        Call GdipDrawImageRect(mlngGraphics, mlngImgHandle, 0, 0, mobjDraw.Width / 15, mobjDraw.Height / 15)

    End Select
    
End Sub

Private Sub ShowNormalImage(ByVal strFileName As String, Optional ByVal sglX As Single = 0, Optional ByVal sglY As Single = 0, Optional ByVal sglW As Single = 0, Optional ByVal sglH As Single = 0, Optional ByVal rfType As RotateFlipType = RotateNoneFlipNone)

    Dim lngGraphics As Long
    Dim lngImgHandle As Long
    Dim W As Long
    Dim H As Long
    Dim lngW As Long, lngH As Long
    Dim sngW As Single, sngH As Single


    On Error GoTo errHandle

    If mblnClear Then mobjDraw.Cls

    Call GdipCreateFromHDC(mobjDraw.hDC, lngGraphics)
    Call GdipLoadImageFromFile(StrPtr(strFileName), lngImgHandle)
    If rfType <> RotateNoneFlipNone Then
        Call GdipImageRotateFlip(lngImgHandle, rfType)
    End If

    Select Case mbytImageStyle
    '------------------------------------------------------------------------------------------------------------------
    Case IMAGESTYLE.自动缩放

        Call GdipGetImageHeight(lngImgHandle, H)    '
        Call GdipGetImageWidth(lngImgHandle, W)

        If W <= mobjDraw.ScaleWidth / 15 And H <= mobjDraw.ScaleHeight / 15 Then
            lngW = W: lngH = H
        Else
            sngW = W / (mobjDraw.ScaleWidth / 15)
            sngH = H / (mobjDraw.ScaleHeight / 15)
            If sngW > sngH Then
                lngW = W / sngW: lngH = H / sngW
            Else
                lngW = W / sngH: lngH = H / sngH
            End If
        End If
        Call GdipDrawImageRect(lngGraphics, lngImgHandle, (mobjDraw.Width / 15 - lngW) / 2, (mobjDraw.Height / 15 - lngH) / 2, lngW, lngH)
    '------------------------------------------------------------------------------------------------------------------
    Case IMAGESTYLE.自动拉伸

        Call GdipDrawImageRect(lngGraphics, lngImgHandle, sglX, sglY, sglW, sglH)
        
    End Select

    mobjDraw.Refresh

    If lngImgHandle > 0 Then
        Call GdipDisposeImage(lngImgHandle)
        lngImgHandle = 0
    End If

    If lngGraphics > 0 Then
        Call GdipDeleteGraphics(lngGraphics)
        lngGraphics = 0
    End If
    
    Exit Sub

    '------------------------------------------------------------------------------------------------------------------
errHandle:
    MsgBox "加载文件错误"
End Sub

Public Sub DanRotate(ByRef picDestHdc As Long, xPos As Long, yPos As Long, ByVal Angle As Long, ByRef picSrcHdc As Long, srcXoffset As Long, srcYoffset As Long, ByVal srcwidth As Long, ByVal srcheight As Long)
    
    Dim Points(3) As POINTAPI
    Dim DefPoints(3) As POINTAPI
    Dim sSin As Single, sCos As Single
    Dim ret As Long

    Points(0).X = -srcwidth * 0.5
    Points(0).Y = -srcheight * 0.5
    Points(1).X = Points(0).X + srcwidth
    Points(1).Y = Points(0).Y
    
    Points(2).X = Points(0).X
    Points(2).Y = Points(0).Y + srcheight
    
    sSin = Sin(Angle * NotPI)
    sCos = Cos(Angle * NotPI)
    DefPoints(0).X = (Points(0).X * sCos - Points(0).Y * sSin) + xPos
    DefPoints(0).Y = (Points(0).X * sSin + Points(0).Y * sCos) + yPos
    
    DefPoints(1).X = (Points(1).X * sCos - Points(1).Y * sSin) + xPos
    DefPoints(1).Y = (Points(1).X * sSin + Points(1).Y * sCos) + yPos
    
    DefPoints(2).X = (Points(2).X * sCos - Points(2).Y * sSin) + xPos
    DefPoints(2).Y = (Points(2).X * sSin + Points(2).Y * sCos) + yPos
    
    Call PlgBlt(picDestHdc, DefPoints(0), picSrcHdc, srcXoffset, srcYoffset, srcwidth, srcheight, 0, 0, 0)
    
'    hdcDest:     指向目标设备环境的句柄?
'
'    lpPoint：      指向代表目标平行四边形3个角的3个顶点的数组指针。源矩形的左上角映射到该数组的第1个顶点，
'                   右上角映射为数组中的第2个顶点，左下角映射成第3个顶点。而右下角则映射成平行四边形中隐含的第4个点。
'
'    hdcSrc:     指向源设备环境的句柄?
'
'    nXSrc：    指定源矩形左上角的X轴坐标，按逻辑单位。
'
'    nYSrc：    指定源矩形左上角的Y轴坐标，按逻辑单位。
'
'    nWidth：   指字源矩形的宽度，按逻辑单位。
'
'    nheight：  指定源矩形的高度，按逻辑单位。
'
'    hbMask:     指向可选的单色位图的句柄?该位图是用来对源矩形的颜色进行屏蔽用的?
'
'    xMask:     指定单色位图左上角的X轴坐标?
'
'    yMask:     指定单色位图左上角的Y轴坐标?


End Sub
