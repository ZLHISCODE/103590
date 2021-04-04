Attribute VB_Name = "mdlDcmProcess"
Option Explicit

'点坐标类型
Public Type TPoint
  X As Integer
  Y As Integer
End Type


Public Sub subSetSharp(objDcmImage As DicomImage, blnSharp As Boolean)
'------------------------------------------------
'功能：dcmView中图像的平滑和锐化
'参数：blnSharp表示图像处理的方向，True=锐化；False=平滑
'返回：无，直接处理dcmView中的图像
'------------------------------------------------
    If Not objDcmImage Is Nothing Then
        If blnSharp = True Then
            '锐化处理
            If objDcmImage.FilterLength <= 0 Then
                objDcmImage.FilterLength = 0
                '先前没有平滑处理，直接进行锐化处理
                objDcmImage.UnsharpEnhancement = objDcmImage.UnsharpEnhancement + 0.1
            Else
                '如果先前已经有平滑处理，则先淡化平滑效果
                objDcmImage.FilterLength = objDcmImage.FilterLength - 1
            End If
        Else
            '平滑处理
            '判断Zoom是否＝1，如果是，则修改为0.9999
            If objDcmImage.ActualZoom = 1 Then
                objDcmImage.Zoom = 0.9999
            End If
            
            If objDcmImage.UnsharpEnhancement <= 0 Then
                objDcmImage.UnsharpEnhancement = 0
                '先前没有锐化处理，直接开始平滑
                '判断FilterLength是否＝0如果是，则在2/ActualZoom和2×FilterLength之间进行调整
                If objDcmImage.FilterLength = 0 Then
                    objDcmImage.FilterLength = 2 / objDcmImage.ActualZoom + 1
                Else    '正常情况下FilterLength＋1
                    objDcmImage.FilterLength = objDcmImage.FilterLength + 1
                End If
            Else
                '先前已经有了锐化处理，先淡化锐化的效果
                objDcmImage.UnsharpEnhancement = objDcmImage.UnsharpEnhancement - 0.1
            End If
        End If
    End If
End Sub



Public Sub subSetRotate(objDcmImage As DicomImage, blnClockwise As Boolean)
'------------------------------------------------
'功能：dcmView中图像的旋转
'参数：blnClockwise旋转的方向,True=顺时针旋转；False=逆时针旋转
'返回：无，直接处理dcmView中的图像
'------------------------------------------------
    If Not objDcmImage Is Nothing Then
        Dim iRotateState As Integer
        
        iRotateState = objDcmImage.RotateState
        If blnClockwise = True Then
            iRotateState = iRotateState - 1
        Else
            iRotateState = iRotateState + 1
        End If
        
        If iRotateState = -1 Then iRotateState = 3
        
        iRotateState = iRotateState Mod 4
        objDcmImage.RotateState = iRotateState
    End If
End Sub


'DicomViewer裁剪后采集图象
Public Function CutImage(objDcmImage As DicomImage) As DicomImage
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
    
    Set CutImage = Nothing
    
    If objDcmImage Is Nothing Then Exit Function
    If objDcmImage.Labels.Count < 1 Then Exit Function
    
    Set img = objDcmImage
    Set lblFrame = objDcmImage.Labels(objDcmImage.Labels.Count)
    
    If Abs(lblFrame.Width) = 0 Or Abs(lblFrame.Height) = 0 Then
        MsgboxCus "请选择图像区域后再保存", vbExclamation, G_STR_HINT_TITLE
        Exit Function
    End If
    
    '图象最大宽高=300
    iMax = 300
    
    '根据label来提取被框选中的图像
    '图象位数,黑白图像为1，彩色图像为3
    iPlane = 1
    If Not IsNull(img.Attributes(&H28, &H4).value) And img.Attributes(&H28, &H4).Exists Then
        If img.Attributes(&H28, &H4).value = "RGB" Then
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
    
    Set CutImage = imgResult
End Function


Public Function GetNewLabel(lType As Integer, lLeft As Integer, lTop As Integer, lWidth As Integer, lHeight As Integer) As DicomLabel
'------------------------------------------------
'功能：生成一个LABEL对象，并对其做初始化。
'参数：lType--标注的类型；lLeft--标注的Left值；lTop--标注的Top值；lWidth--标注的Width值；lHeight--标注的Height值。
'返回：新生成的标注。
'编制人：黄捷
'------------------------------------------------
    Dim l As New DicomLabel
    
    l.LabelType = lType
    l.XOR = True
    l.ImageTied = True
    l.Left = lLeft
    l.Top = lTop
    l.Width = lWidth
    l.Height = lHeight
    l.Margin = 0
    l.AutoSize = True
    l.FontSize = 12
    l.LineWidth = 1
    
    If l.LabelType = 0 Then     '文字
        l.Transparent = False
        l.Width = 200
        l.Height = 10
    End If
    
    Set GetNewLabel = l
End Function
   
   
Public Sub subCenterZoom(frmWindow As Form, img As DicomImage, Viewer As DicomViewer, dblZoom As Double, corpSize As TPoint)
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
            
    img.ScrollX = (img.SizeX * img.ActualZoom - frmWindow.ScaleX(Viewer.Width, vbTwips, vbPixels) / Viewer.MultiColumns) / 2 + corpSize.X
    img.ScrollY = (img.SizeY * img.ActualZoom - frmWindow.ScaleY(Viewer.Height, vbTwips, vbPixels) / Viewer.MultiRows) / 2 + corpSize.Y
End Sub


Public Sub RectangleZoom(Viewer As DicomViewer, img As DicomImage, lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long)
    Dim newZoom As Double
    Dim dblRatio As Double
    Dim sX As Long
    Dim sY As Long
    Dim oldZoom As Double
    
    If lngWidth > 0 And lngHeight > 0 Then
        oldZoom = img.ActualZoom
        sX = img.ActualScrollX
        sY = img.ActualScrollY
        
        img.StretchToFit = False
        
        dblRatio = Viewer.Width / Screen.TwipsPerPixelX / lngWidth
        If dblRatio > Viewer.Height / Screen.TwipsPerPixelY / lngHeight Then
            dblRatio = Viewer.Height / Screen.TwipsPerPixelY / lngHeight
        End If
        
        newZoom = oldZoom * dblRatio
        img.Zoom = newZoom
        
        img.ScrollX = (sX + lngLeft) * dblRatio
        img.ScrollY = (sY + lngTop) * dblRatio
    End If
End Sub
