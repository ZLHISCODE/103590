Attribute VB_Name = "mdlBarCode2D"
Option Explicit

Public Function DrawBarCode2D(ByVal strText As String, picTemp As PictureBox, Optional lngSize As Long) As StdPicture
'功能：返回QR二维条码图片
'参数：lngSize=返回以Twip为单位的图片合适尺寸
'返回：QR二维条码图片，可能是放大了的BMP图片
    Static objQRMaker As Object
    Static intInstall As Integer '0-未检测,1-已安装,-1-未安装
    
    Dim strFile As String
    Dim objPic As StdPicture
    
    If intInstall = 0 Then
        On Local Error Resume Next
        Set objQRMaker = CreateObject("QRMAKER.QRmakerCtrl.1")
        Err.Clear: On Local Error GoTo 0
        intInstall = IIF(objQRMaker Is Nothing, -1, 1)
        
        '初始化控件属性
        If intInstall = 1 Then
            With objQRMaker
                .GapAdjust = 0 'GpAjOff
                .LanguageCode = 1
                
                .EccLevel = 1 'M
                .ModelNo = 2 'Model2
                .Rotate = 0 'D0
                
                .QuietZone = 1
                
                .ForeWColor = vbWhite
                .ForeBColor = vbBlack
            End With
        End If
    End If
    
    lngSize = 0
    
    If intInstall = -1 Then
        picTemp.AutoRedraw = True
        picTemp.BorderStyle = 0
        picTemp.ScaleMode = vbTwips
        picTemp.Cls
        
        lngSize = picTemp.ScaleX(50, vbPixels, vbTwips)
        picTemp.Width = lngSize: picTemp.Height = lngSize
        
        picTemp.DrawWidth = 1
        picTemp.Line (0, 0)-(picTemp.Width, picTemp.Height), vbBlack
        picTemp.Line (picTemp.Width, 0)-(0, picTemp.Height), vbBlack
        picTemp.DrawWidth = 2
        picTemp.Line (0, 0)-(picTemp.Width, picTemp.Height), vbBlack, B
        
        Set DrawBarCode2D = picTemp.Image
        picTemp.Cls
    Else
        If strText = "" Then strText = "重庆中联信息产业有限公司"
        strFile = gobjFile.GetSpecialFolder(TemporaryFolder) & "\" & gobjFile.GetTempName
        
        objQRMaker.InputData = strText
        objQRMaker.CreateQrMetaFile 0, strFile, 2 'BMP
        
        Set objPic = LoadPicture(strFile)
        gobjFile.DeleteFile strFile, True
        
        picTemp.ScaleMode = vbTwips
        lngSize = picTemp.ScaleX(2 * (objQRMaker.NumCell + objQRMaker.QuietZone * 2), vbPixels, vbTwips)
        
        Set DrawBarCode2D = objPic
    End If
End Function
