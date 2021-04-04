VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form frmLisChart 
   BorderStyle     =   0  'None
   Caption         =   "Chart"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin C1Chart2D8.Chart2D ChartThis 
      Height          =   2760
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2760
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   4868
      _ExtentY        =   4868
      _StockProps     =   0
      ControlProperties=   "frmLisChart.frx":0000
   End
End
Attribute VB_Name = "frmLisChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mcurYunti As Currency 'X坐标的一个点相当于多少个Y坐标的点
Private Xpixe As Currency, Ypixe As Currency

Private Sub Form_Activate()
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.Hide
End Sub

Public Function DrawImg(ByVal strType As String, ByVal strData As String, ByVal strFileName As String, Optional intSaveType As Integer = 1) As Boolean
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '功能:          画图现支持(RBC,PLT,BAS,WBC)
        '参数:          strType  图像名称
        '                   strData  图像数据
        '                   strFileName 生成的图片文件名
        '                   intSaveType 0-cht格式 1-jpg格式 2-png格式
        '
        '其他               数据的第一位 0=直方图 1=散点图 2=血流变粘度特征曲线图 3=血沉曲线图
        '                   100=图片
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim aImage() As String
        Dim aFormat() As String          '显示格式
        Dim aLines() As String
        Dim lngLoop As Long
        Dim lngLoopX As Long
        Dim i As Integer, j As Integer
        Dim dblX As Double, dblY As Double
        Dim aLables() As String, aLable() As String, strLable As String
        Dim strFile As String, strImgData As String
        Dim killFile As String
        Dim FrmObj As frmLisChartPic
        
        On Error GoTo errH
    
     aImage = Split(strData, ";")
 
     If aImage(0) = 0 Then
         '直方图 clsLISDev_ABX_P120
         With Me.ChartThis
             .IsBatched = True
             .Reset
             .ChartGroups(1).Data.NumSeries = 0
             .Header.Adjust = oc2dAdjustCenter
             .Header.Text = strType
             .Header.Font.Bold = True
             .Header.Font.Size = 12
             .ChartGroups(1).Styles(1).Line.Color = vbBlack
             .ChartGroups(1).Styles(1).Line.Width = 1
             .ChartGroups(1).Styles(1).Symbol.Shape = oc2dShapeNone
         
             .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
             .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
         
             .ChartArea.Axes("X").ValueLabels.RemoveAll
     
             aFormat = Split(aImage(1), ",")
             If UBound(aFormat) > 1 Then
                 .ChartArea.Axes("Y").Min = aFormat(0)
                 .ChartArea.Axes("X").Max = aFormat(1)
                 '----2008-03-28 有时不显示图形
                 .ChartArea.Axes("Y").Origin = 0
                 .ChartArea.Axes("Y").Min = 0
             
                 .ChartArea.Axes("Y").Max.IsDefault = True
                 .ChartArea.Axes("X").Max.IsDefault = True
                 .ChartArea.Axes("Y").Min.IsDefault = True
                 .ChartArea.Axes("X").Min.IsDefault = True
             
                 For i = 2 To UBound(aFormat)
                     .ChartArea.Axes("X").ValueLabels.Add Mid(aFormat(i), 1, InStr(aFormat(i), "-") - 1), Mid(aFormat(i), InStr(aFormat(i), "-") + 1)
                 Next
             End If
             .ChartGroups(1).Data.NumSeries = 1
             .ChartGroups(1).Data.NumPoints(1) = UBound(aImage)
             For i = 2 To UBound(aImage)
                 .ChartGroups(1).Data.y(1, i) = Val(aImage(i))
             Next
             .IsBatched = False
             If intSaveType = 1 Then
                 DrawImg = .SaveImageAsJpeg(strFileName, 100, False, False, False)
             ElseIf intSaveType = 2 Then
                 DrawImg = .SaveImageAsPng(strFileName, False)
             Else
                 DrawImg = .Save(strFileName)
             End If
         End With
     
     End If
 
     '-- 散点图 clsLISDev_ABX_P120
     If aImage(0) = 1 Then
         With Me.ChartThis
             .IsBatched = True
             .Reset
             .ChartGroups(1).Data.NumSeries = 0
             .Header.Adjust = oc2dAdjustCenter
             .Header.Text = strType
             .Header.Font.Bold = True
             .Header.Font.Size = 12
             .ChartArea.PlotArea.IsBoxed = True
             .ChartGroups(1).Data.NumSeries = UBound(aImage) - 1
         
             .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
             .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
             For lngLoop = UBound(aImage) - 1 To 1 Step -1
                 .ChartGroups(1).ChartType = oc2dTypePlot
                 .ChartGroups(1).Styles(lngLoop).Line.Pattern = oc2dLineNone
                 .ChartGroups(1).Styles(lngLoop).Line.Color = vbBlack
                 .ChartGroups(1).Styles(lngLoop).Symbol.Shape = oc2dShapeBox
                 .ChartGroups(1).Styles(lngLoop).Symbol.Size = 1
                 .ChartGroups(1).Styles(lngLoop).Symbol.Color = vbBlack
                 .ChartGroups(1).Data.NumPoints(lngLoop) = Len(aImage(lngLoop)) + 1
                 For lngLoopX = 1 To Len(aImage(lngLoop)) + 1
                     .ChartGroups(1).Data.y(lngLoop, lngLoopX) = IIf(Mid(aImage(lngLoop), lngLoopX, 1) = 0, .ChartGroups(1).Data.HoleValue, 128 - lngLoop + 1)
                 Next
             Next
             .IsBatched = False
             If intSaveType = 1 Then
                 DrawImg = .SaveImageAsJpeg(strFileName, 100, False, False, False)
             ElseIf intSaveType = 2 Then
                 DrawImg = .SaveImageAsPng(strFileName, False)
             Else
                 DrawImg = .Save(strFileName)
             End If
         End With
     End If

     '--- 血流变图  clsLISDev_File_LBYN6C
     If aImage(0) = 2 Then
         DrawImg = ChartDraw血流变(strType, aImage(1), aImage(2), aImage(3), strFileName, intSaveType)
     End If
 
     '--血沉曲线图 clsLISDev_File_LBYN6C
     If aImage(0) = 3 Then
         DrawImg = ChartDraw血沉(strType, aImage(1), aImage(2), aImage(3), strFileName, intSaveType)
     End If
 
     '--- 有两根重叠曲线的PLT图 clsLISDev_HMX
     If aImage(0) = 4 Then
         DrawImg = ChartDrawPLT(strType, aImage(1), aImage(2), strFileName, intSaveType)
     End If
 
     '--- 在本地的PIC控件上绘制 直方图 然后显示
     If aImage(0) = 5 Then
         DrawImg = PicShowChart(strType & ";" & strData, strFileName, intSaveType)
     
     End If
 
     '--- 直方图·在同一个图上绘制多条曲线
     If aImage(0) = 6 Then
         '直方图 WBC clsLISDev_MEDONIC_M20M
         With Me.ChartThis
             .IsBatched = True
             .Reset
             .ChartGroups(1).Data.NumSeries = 0
             .Header.Adjust = oc2dAdjustCenter
             .Header.Text = strType
             .Header.Font.Bold = True
             .Header.Font.Size = 12
         
         
             .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
             .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
         
             .ChartArea.Axes("X").ValueLabels.RemoveAll
             aFormat = Split(aImage(1), ",")
             If UBound(aFormat) > 1 Then
                 .ChartArea.Axes("Y").Min = aFormat(0)
                 .ChartArea.Axes("X").Max = aFormat(1)
                 '----2008-03-28 有时不显示图形
                 .ChartArea.Axes("Y").Origin = 0
                 .ChartArea.Axes("Y").Min = 0
             
                 For i = 2 To UBound(aFormat)
                     .ChartArea.Axes("X").ValueLabels.Add Mid(aFormat(i), 1, InStr(aFormat(i), "-") - 1), Mid(aFormat(i), InStr(aFormat(i), "-") + 1)
                 Next
             End If
         
             aLines = Split(strData, "~")
             .ChartGroups(1).Data.NumSeries = UBound(aLines)
         
             For i = LBound(aLines) + 1 To UBound(aLines)
                 aImage = Split(aLines(i), ";")
                 .ChartGroups(1).Styles(i).Line.Color = vbBlack
                 .ChartGroups(1).Styles(i).Line.Width = 1
                 .ChartGroups(1).Styles(i).Symbol.Shape = oc2dShapeNone
                 .ChartGroups(1).Data.NumPoints(i) = UBound(aImage) + 1
                 For j = LBound(aImage) To UBound(aImage)
                     .ChartGroups(1).Data.y(i, j + 1) = Val(aImage(j))
                 Next
             Next
             .IsBatched = False
             If intSaveType = 1 Then
                 DrawImg = .SaveImageAsJpeg(strFileName, 100, False, False, False)
             ElseIf intSaveType = 2 Then
                 DrawImg = .SaveImageAsPng(strFileName, False)
             Else
                 DrawImg = .Save(strFileName)
             End If
         End With
     End If
     killFile = ""
     
    If aImage(0) >= 100 And aImage(0) <= 227 Then
        strFile = aImage(1)
             
        If UCase$(strFile) Like "*.ZIP" Then
            killFile = strFile
            If aImage(0) >= 200 And aImage(0) <= 207 Then
                strFile = zlFileUnzip(strFile)
            ElseIf aImage(0) >= 210 And aImage(0) <= 217 Then
                strFile = zlFileUnzip(strFile)
            ElseIf aImage(0) >= 220 And aImage(0) <= 227 Then
                strFile = zlFileUnzip(strFile)
            End If
            
            If killFile <> "" Then Kill killFile: killFile = "" '解压后的原始ZIP要删除
        End If
         
            '要先保存成bmp才可以作为chart2D的背景...
            Set FrmObj = New frmLisChartPic
            If UCase(strFile) Like "*.JPG" Then
                FrmObj.picTmp.Picture = LoadPicture(strFile)
                strFile = Replace(UCase(strFile), ".JPG", ".BMP")
                SavePicture FrmObj.picTmp, strFile
                killFile = Replace(UCase(strFile), ".BMP", ".JPG")
            ElseIf UCase(strFile) Like "*.GIF" Then
                If CheckGif(strFile) Then
                    FrmObj.picTmp.Picture = LoadPicture(strFile)
                    strFile = Replace(UCase(strFile), ".GIF", ".BMP")
                    SavePicture FrmObj.picTmp, strFile
                Else
                    Exit Function
                End If
            End If
            Unload FrmObj
            
         '--- 直接显示图片 clsLISDev_UF100_DY
            If CInt(Val(Right(aImage(0), 1))) = 0 Then
                DrawImg = ChartShowPic(strType, strFile, strFileName, , intSaveType)  '用默认的layOut
            Else
                DrawImg = ChartShowPic(strType, strFile, strFileName, CInt(Val(Right(aImage(0), 1))), intSaveType)   '用指定的layout
            End If
            If strFile <> "" Then Kill strFile
            If killFile <> "" Then Kill killFile '删除临时文件
    End If
    
    Exit Function
errH:
    MsgBox err.Description, vbExclamation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function ChartDraw血流变(ByVal strType As String, ByVal strXYin As String, ByVal strLineIn As String, ByVal strLableIn As String, ByVal strFileName As String, Optional intSaveType As Integer) As Boolean
    Dim aFormat() As String
    Dim aLables() As String, aLable() As String, strLable As String
    Dim i As Integer
    Dim aCurves() As String '存曲线数据
    Dim aCurve() As String
    Dim intPoint As Integer
    Dim aPoint() As String '存描点数据
    Dim lngLoop As Long
    Dim dblX As Double, dblY As Double
    Dim aAxes() As String
    
    With Me.ChartThis
        .IsBatched = True
        .Reset
        .ChartGroups(1).Data.NumSeries = 0
        .Header.Adjust = oc2dAdjustCenter
        .Header.Text = strType
        .Header.Font.Bold = True
        .Header.Font.Size = 12
        
        .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
        .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
        
        .ChartArea.Axes("X").ValueLabels.RemoveAll
        .ChartArea.Axes("Y").ValueLabels.RemoveAll
        
        '--- 开始
        aFormat = Split(strXYin, "|") '画坐标
        If UBound(aFormat) > 1 Then
            
            aAxes = Split(aFormat(0), ",") '坐标长度
            .ChartArea.Axes("Y").Max = aAxes(0)
            .ChartArea.Axes("X").Max = aAxes(1)
            .ChartArea.Axes("Y").Origin = 0
            
            aAxes = Split(aFormat(1), ",") 'X刻度
            For i = LBound(aAxes) To UBound(aAxes)
                .ChartArea.Axes("X").ValueLabels.Add Mid(aAxes(i), 1, InStr(aAxes(i), "-") - 1), Mid(aAxes(i), InStr(aAxes(i), "-") + 1)
            Next
            
            aAxes = Split(aFormat(2), ",") 'Y刻度
            For i = LBound(aAxes) To UBound(aAxes)
                .ChartArea.Axes("Y").ValueLabels.Add Mid(aAxes(i), 1, InStr(aAxes(i), "-") - 1), Mid(aAxes(i), InStr(aAxes(i), "-") + 1)
            Next
            
        End If
        
        aFormat = Split(strLineIn, "~") '曲线及描点
        If UBound(aFormat) > 0 Then

            
            aCurves = Split(aFormat(0), "|") '曲线数据
            
            For i = LBound(aCurves) To UBound(aCurves)
                aCurve = Split(aCurves(i), ",")
                
                .ChartGroups(1).Styles(i + 1).Line.Color = vbBlack
                .ChartGroups(1).Styles(i + 1).Line.Width = 1
                .ChartGroups(1).Styles(i + 1).Symbol.Shape = oc2dShapeNone
                
                .ChartGroups(1).Data.NumSeries = i + 1
                .ChartGroups(1).Data.NumPoints(i + 1) = 225
            
                    '画线
                If UBound(aCurve) > 2 Then
                    For lngLoop = 1 To 200
                        dblY = GetNd(Val(aCurve(0)), Val(aCurve(1)), Val(aCurve(2)), Val(aCurve(3)), lngLoop)
                        If lngLoop > 2 Then
                            .ChartGroups(1).Data.y(i + 1, lngLoop + 1) = dblY
                        Else
                            .ChartGroups(1).Data.y(i + 1, lngLoop + 1) = .ChartGroups(1).Data.HoleValue
                        End If
                    Next
                End If
            Next
        
            aPoint = Split(aFormat(1), ",") '描点数据

            intPoint = UBound(aCurves) + 2
            .ChartGroups(1).Styles(intPoint).Line.Pattern = oc2dLineNone
            .ChartGroups(1).Styles(intPoint).Line.Color = vbBlack
            .ChartGroups(1).Styles(intPoint).Line.Width = 1
            .ChartGroups(1).Styles(intPoint).Symbol.Color = vbBlack
            .ChartGroups(1).Styles(intPoint).Symbol.Shape = oc2dShapeSquare
            .ChartGroups(1).Data.NumSeries = intPoint
            .ChartGroups(1).Data.NumPoints(intPoint) = 225

            For i = 1 To 200
                
                For lngLoop = LBound(aPoint) To UBound(aPoint)
                    '-- 描点
                    dblX = Val(Mid(aPoint(lngLoop), 1, InStr(aPoint(lngLoop), "-") - 1))
                    dblY = Val(Mid(aPoint(lngLoop), InStr(aPoint(lngLoop), "-") + 1))
                    If dblX = i + 1 Then
                        .ChartGroups(1).Data.y(intPoint, i + 1) = dblY
                        Exit For
                    Else
                        .ChartGroups(1).Data.y(intPoint, i + 1) = .ChartGroups(1).Data.HoleValue
                    End If
                Next
                
            Next
        End If
        
        
        aLables = Split(strLableIn, "~")  'X轴标签，Y轴标签
        
        aLable = Split(aLables(0), ",")
        strLable = aLable(0)
        If strLable <> "" Then
            .ChartLabels.Add
            .ChartLabels.Item(1).Text = strLable
            .ChartLabels.Item(1).AttachDataCoord.x = Val(aLable(1))
            .ChartLabels.Item(1).AttachDataCoord.y = Val(aLable(2))
        End If
        
        aLable = Split(aLables(1), ",")
        strLable = aLable(0)
        If strLable <> "" Then
            .ChartLabels.Add
            .ChartLabels.Item(2).Text = strLable
            .ChartLabels.Item(2).AttachDataCoord.x = Val(aLable(1))
            .ChartLabels.Item(2).AttachDataCoord.y = Val(aLable(2))
        End If
        
        '---- 结束

        .IsBatched = False
        If intSaveType = 1 Then
            ChartDraw血流变 = .SaveImageAsJpeg(strFileName, 100, False, False, False)
        ElseIf intSaveType = 2 Then
            ChartDraw血流变 = .SaveImageAsPng(strFileName, False)
        Else
            ChartDraw血流变 = .Save(strFileName)
        End If
    End With
End Function
Private Function ChartDraw血沉(ByVal strType As String, ByVal strXYin As String, ByVal strLineIn As String, ByVal strLableIn As String, ByVal strFileName As String, Optional intSaveType As Integer) As Boolean
    '画血沉图
    Dim aFormat() As String
    Dim aLables() As String, aLable() As String, strLable As String
    Dim i As Integer
    Dim aCurves() As String '存曲线数据
    Dim aCurve() As String
    Dim intPoint As Integer
    Dim aPoint() As String '存描点数据
    Dim lngLoop As Long
    Dim dblX As Double, dblY As Double
    Dim aAxes() As String
    
    With Me.ChartThis
        .IsBatched = True
        .Reset
        .ChartGroups(1).Data.NumSeries = 0
        .Header.Adjust = oc2dAdjustCenter
        .Header.Text = strType
        .Header.Font.Bold = True
        .Header.Font.Size = 12
        
        .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
        .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
        
        .ChartArea.Axes("X").ValueLabels.RemoveAll
        .ChartArea.Axes("Y").ValueLabels.RemoveAll
        
        .ChartArea.Axes("Y").Origin = 0
        .ChartArea.Axes("Y").Min = 0
        .ChartArea.Axes("X").Origin = 0
        .ChartArea.Axes("X").Min = 0
        '--- 开始
        aFormat = Split(strXYin, "|") '画坐标
        If UBound(aFormat) > 1 Then
                            
            aAxes = Split(aFormat(0), ",") '坐标长度
            .ChartArea.Axes("Y").Max = aAxes(0)
            .ChartArea.Axes("X").Max = aAxes(1)
            .ChartArea.Axes("Y").Origin = 0
            
            aAxes = Split(aFormat(1), ",") 'X刻度
            For i = LBound(aAxes) To UBound(aAxes)
                .ChartArea.Axes("X").ValueLabels.Add Mid(aAxes(i), 1, InStr(aAxes(i), "-") - 1), Mid(aAxes(i), InStr(aAxes(i), "-") + 1)
            Next
            
            aAxes = Split(aFormat(2), ",") 'Y刻度
            For i = LBound(aAxes) To UBound(aAxes)
                .ChartArea.Axes("Y").ValueLabels.Add Mid(aAxes(i), 1, InStr(aAxes(i), "-") - 1), Mid(aAxes(i), InStr(aAxes(i), "-") + 1)
            Next
            
        End If
        
        aFormat = Split(strLineIn, ",") '曲线
        If UBound(aFormat) > 0 Then

            .ChartGroups(1).Styles(1).Line.Color = vbBlack
            .ChartGroups(1).Styles(1).Line.Width = 1
            .ChartGroups(1).Styles(1).Symbol.Shape = oc2dShapeSquare '空心方点
            .ChartGroups(1).Styles(1).Symbol.Color = vbRed
            
            .ChartGroups(1).Styles(2).Symbol.Color = vbRed
            .ChartGroups(1).Styles(2).Symbol.Shape = oc2dShapeSquare
            .ChartGroups(1).Styles(2).Line.Color = vbBlack
            
            .ChartGroups(1).Styles(3).Symbol.Shape = oc2dShapeNone
            .ChartGroups(1).Styles(3).Line.Color = vbBlack
            
            .ChartGroups(1).Data.NumSeries = 3
            .ChartGroups(1).Data.NumPoints(1) = UBound(aFormat)
            
'            For i = 1 To UBound(aFormat) * 4
'                    '画线
'                If (i Mod 2) = 0 Then
'                    .ChartGroups(1).Data.Y(1, i) = aFormat(i / 4)
'                End If
'                If aFormat(i / 4) - 0.5 >= 0 Then
'                .ChartGroups(1).Data.Y(2, i) = aFormat(i / 4) - 0.5
'                End If
'            Next
            For i = 1 To UBound(aFormat)
                If (i Mod 2) = 0 Then
                    .ChartGroups(1).Data.y(1, i) = aFormat(i)
                Else
                    .ChartGroups(1).Data.y(2, i) = aFormat(i)
                End If
                    '画线
                .ChartGroups(1).Data.y(3, i) = aFormat(i) - 0.3
                
            Next
        
        End If
        
        aLables = Split(strLableIn, "~") 'X轴标签，Y轴标签
        
        aLable = Split(aLables(0), ",")
        strLable = aLable(0)
        If strLable <> "" Then
            .ChartLabels.Add
            .ChartLabels.Item(1).Text = strLable
            .ChartLabels.Item(1).AttachDataCoord.x = Val(aLable(1))
            .ChartLabels.Item(1).AttachDataCoord.y = Val(aLable(2))
        End If
        
        aLable = Split(aLables(1), ",")
        strLable = aLable(0)
        If strLable <> "" Then
            .ChartLabels.Add
            .ChartLabels.Item(2).Text = strLable
            .ChartLabels.Item(2).AttachDataCoord.x = Val(aLable(1))
            .ChartLabels.Item(2).AttachDataCoord.y = Val(aLable(2))
        End If
        
        '---- 结束

        .IsBatched = False

        If intSaveType = 1 Then
            ChartDraw血沉 = .SaveImageAsJpeg(strFileName, 100, False, False, False)
        ElseIf intSaveType = 2 Then
            ChartDraw血沉 = .SaveImageAsPng(strFileName, False)
        Else
            ChartDraw血沉 = .Save(strFileName)
        End If
    End With
End Function
Private Function ChartShowPic(ByVal strType As String, ByVal strImgName As String, ByVal strFileName As String, Optional ByVal intLayOut As Integer = oc2dImageFitted, Optional intSaveType As Integer = 0) As Boolean
    'Chart 显示图片
    'strImgName-源文件 strFilename-生成文件名
    Dim strImgFile As String
    
    With Me.ChartThis
        .IsBatched = True
        .Reset
        .ChartGroups(1).Data.NumSeries = 0
        .Header.Adjust = oc2dAdjustCenter
        .Header.Text = strType
        .Header.Font.Bold = True
        .Header.Font.Size = 12
        
        .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
        .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
        
        .ChartArea.Axes("X").ValueLabels.RemoveAll
        .ChartArea.Axes("Y").ValueLabels.RemoveAll
        If intSaveType = 0 Then
            strImgFile = Replace(strFileName, ".cht", ".bmp")
            gobjFile.CopyFile strImgName, strImgFile, True
        Else
            strImgFile = strFileName
            gobjFile.CopyFile strImgName, strImgFile, True
        End If
        .Interior.Image.Filename = strImgName
        .Interior.Image.Layout = intLayOut 'oc2dImageFitted
'        .ChartArea.Interior.Image.Filename = strImgName
'        .ChartArea.Interior.Image.Layout = intLayOut
        .IsBatched = False
        If intSaveType = 1 Then
            ChartShowPic = .SaveImageAsJpeg(strFileName, 100, False, False, False)
        ElseIf intSaveType = 2 Then
            ChartShowPic = .SaveImageAsPng(strFileName, False)
        Else
            ChartShowPic = .Save(strFileName)
        End If
    End With
    
End Function
Private Function GetNd(ByVal ND0 As Double, QB0 As Double, ND1 As Double, QB1 As Double, ByVal Qb As Double) As Double
    '血流变标准曲线坐标计算函数
    Dim k0 As Double, K1 As Double
    Dim sn As Double

    k0 = (Sqr(ND0) - Sqr(ND1)) / (1 / (Sqr(QB0)) - 1 / (Sqr(QB1)))
    K1 = Sqr(ND0) - k0 * (1 / (Sqr(QB0)))
    sn = k0 * (1 / (Sqr(Qb))) + K1
    GetNd = sn * sn

End Function

Private Function ChartDrawPLT(ByVal strType As String, ByVal str_座标 As String, ByVal str_Lines As String, ByVal strFileName As String, Optional intSaveType As Integer = 0) As Boolean
    
    Dim aFormat() As String
    Dim i As Integer, y As Integer
    Dim varLine() As String
    Dim aLine() As String
    Dim Y轴 As String, str轴标题 As String
    Dim aLables() As String, aLable() As String, strLable As String
    Y轴 = "": str轴标题 = ""
     '2根线重叠绘制的直方图
    With Me.ChartThis
        .IsBatched = True
        .Reset
        .ChartGroups(1).Data.NumSeries = 0
        .Header.Adjust = oc2dAdjustCenter
        .Header.Text = strType
        .Header.Font.Bold = True
        .Header.Font.Size = 12
        .ChartGroups(1).Styles(1).Line.Color = vbBlack
        .ChartGroups(1).Styles(1).Line.Width = 1
        .ChartGroups(1).Styles(1).Symbol.Shape = oc2dShapeNone
        
        .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
        .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
        
        .ChartArea.Axes("X").ValueLabels.RemoveAll
        
        If InStr(str_座标, "|") > 0 Then
            Y轴 = Split(str_座标, "|")(1)
            str_座标 = Split(str_座标, "|")(0)
        End If
        aFormat = Split(str_座标, ",")
        If UBound(aFormat) > 1 Then
            .ChartArea.Axes("Y").Max = aFormat(0)
            .ChartArea.Axes("X").Max = aFormat(1)
            
            .ChartArea.Axes("Y").Min = 0
            .ChartArea.Axes("X").Min = 0
           
           .ChartArea.Axes("Y").Origin = 0
           .ChartArea.Axes("X").Origin = 0
            .ChartArea.Axes("Y").Max.IsDefault = True
            .ChartArea.Axes("X").Max.IsDefault = True
            .ChartArea.Axes("Y").Min.IsDefault = True
            .ChartArea.Axes("X").Min.IsDefault = True
            For i = 2 To UBound(aFormat)
                .ChartArea.Axes("X").ValueLabels.Add Mid(aFormat(i), 1, InStr(aFormat(i), "-") - 1), Mid(aFormat(i), InStr(aFormat(i), "-") + 1)
            Next
        End If
        If Y轴 <> "" Then
            aFormat = Split(Y轴, ",")
            For i = 0 To UBound(aFormat)
                .ChartArea.Axes("Y").ValueLabels.Add Mid(aFormat(i), 1, InStr(aFormat(i), "-") - 1), Mid(aFormat(i), InStr(aFormat(i), "-") + 1)
            Next
        End If
        
        If InStr(str_Lines, "~") > 0 Then
            str轴标题 = Split(str_Lines, "~")(1)
            str_Lines = Split(str_Lines, "~")(0)
        End If
        varLine = Split(str_Lines, "|")
        For i = 0 To UBound(varLine)
            aLine = Split(varLine(i), ",")
            .ChartGroups(1).Data.NumSeries = i + 1
            .ChartGroups(1).Data.NumPoints(i + 1) = UBound(aLine) + 1
            
            .ChartGroups(1).Styles(i + 1).Line.Color = vbBlack
            .ChartGroups(1).Styles(i + 1).Line.Width = 1
            .ChartGroups(1).Styles(i + 1).Symbol.Shape = oc2dShapeNone
            
            For y = 0 To UBound(aLine)
                .ChartGroups(1).Data.y(i + 1, y + 1) = IIf(Val(aLine(y)) >= 1, aLine(y), .ChartGroups(1).Data.HoleValue)
            Next
        Next
        
        If str轴标题 <> "" Then
            aLables = Split(str轴标题, "|") 'X轴标签，Y轴标签
            
            aLable = Split(aLables(0), ",")
            strLable = aLable(0)
            If strLable <> "" Then
                .ChartLabels.Add
                .ChartLabels.Item(1).Text = strLable
                .ChartLabels.Item(1).AttachDataCoord.x = Val(aLable(1))
                .ChartLabels.Item(1).AttachDataCoord.y = Val(aLable(2))
            End If
            
            aLable = Split(aLables(1), ",")
            strLable = aLable(0)
            If strLable <> "" Then
                .ChartLabels.Add
                .ChartLabels.Item(2).Text = strLable
                .ChartLabels.Item(2).AttachDataCoord.x = Val(aLable(1))
                .ChartLabels.Item(2).AttachDataCoord.y = Val(aLable(2))
            End If
        End If
        .IsBatched = False
    
        If intSaveType = 1 Then
            ChartDrawPLT = .SaveImageAsJpeg(strFileName, 100, False, False, False)
        ElseIf intSaveType = 2 Then
            ChartDrawPLT = .SaveImageAsPng(strFileName, False)
        Else
            ChartDrawPLT = .Save(strFileName)
        End If
    End With

End Function


'------
Private Function PicShowChart(ByVal strData As String, strFileName As String, Optional intSaveType As Integer = 0) As Boolean
    '在PIC上画图然后转存到Chart控件上显示
    Dim strImgName As String
    Dim frmPic As New frmLisGraph
    
    If strData <> "" Then
        With frmPic
            If DrawGam(.Picture1, strData) Then
                strImgName = strFileName
                If InStr(strImgName, ".") > 0 Then
                    strImgName = Mid(strImgName, 1, InStr(strImgName, ".")) & "Bmp"
                    If Dir(strImgName) <> "" Then Kill strImgName
                    SavePicture .Picture1.Image, strImgName
                    Call ChartShowPic("", strImgName, strFileName, 5, intSaveType)
                    PicShowChart = True
                    'If mobjFSO.FileExists(strImgName) Then mobjFSO.DeleteFile strImgName
                End If
            End If
        End With
    
    End If
    If Dir(strImgName) <> "" Then Kill strImgName
    Unload frmPic
End Function

Private Function DrawGam(ByRef objPic As PictureBox, ByVal strData As String) As Boolean
        '数据格式:
        '标题;图像类型;Y高度,X长度;上下左右边框;X轴刻度[|Y刻度];曲线数据[;界标数据]
        ' 其中:曲线数据: 是y座标数据,以,分隔,多条曲线数据以|分隔
        '     :界标数据: 是x座标数据,以,号分隔
        Dim curX As Currency, curY As Currency, curLastY As Currency
        Dim intWidth As Integer, intHight As Integer
        Dim str上下左右边框 As String
        Dim int边框左 As Integer, int边框右 As Integer, int边框上 As Integer, int边框下 As Integer
        Dim lngColor As Long
    
        Dim varData As Variant
        Dim str坐标 As String, str标题 As String, strLineData As String, str刻度 As String
        Dim var坐标  As Variant, strK As String, curK As Currency, var刻度 As Variant
        Dim str界标 As String
        Dim varLindData As Variant
        Dim lngLoop As Long
        Dim curOldW  As Currency, curOldH As Currency, curOldSW As Currency, curOldSH As Currency
        On Error GoTo errHandle
    
     varData = Split(strData, ";")
 
     If UBound(varData) < 3 Then Exit Function '---数据不全
 
     If varData(1) <> "5" Then Exit Function '---格式不符
     str标题 = varData(0)
     str坐标 = varData(2)
     str上下左右边框 = varData(3)
     str刻度 = varData(4)
     strLineData = varData(5)
     varLindData = Split(strLineData, "|")
 
     If UBound(varData) > 5 Then
         str界标 = varData(6)
     End If
 
     '定义大小
     var坐标 = Split(str坐标, ",")
 
     intHight = var坐标(0): intWidth = var坐标(1)
     mcurYunti = intHight / intWidth
     If str上下左右边框 = "" Then
         int边框左 = 20: int边框右 = 10: int边框上 = 10 * mcurYunti: int边框下 = 50 * mcurYunti
     Else

         int边框上 = Split(str上下左右边框, ",")(0) * mcurYunti
         int边框下 = Split(str上下左右边框, ",")(1) * mcurYunti
         int边框左 = Split(str上下左右边框, ",")(2)
         int边框右 = Split(str上下左右边框, ",")(3)
     End If
 
     objPic.Cls
     objPic.BackColor = vbWhite
     curOldW = objPic.Width
     curOldH = objPic.Height
 
     objPic.Width = 3000
     objPic.Height = 1500
     objPic.DrawMode = vbCopyPen '缺省 画笔
     objPic.DrawStyle = vbSolid  'VbSolid -实线 VbDash-虚线
     objPic.DrawWidth = 1.5        '线宽
     objPic.AutoRedraw = True
 
     'objpic.Height = objpic.Width * (intHight / intWidth)
 
     Dim curTmp As Currency
     curOldSW = objPic.ScaleWidth
     curOldSH = objPic.ScaleHeight
 
     curTmp = objPic.ScaleWidth / (intWidth + int边框左 + int边框右)
     Xpixe = curTmp / Screen.TwipsPerPixelX  '现在一个X点=多少像素
 
     curTmp = objPic.ScaleHeight / (intHight + int边框上 + int边框下)
     Ypixe = curTmp / Screen.TwipsPerPixelY
 
     objPic.Scale (0, 0)-(intWidth + int边框左 + int边框右, intHight + int边框上 + int边框下)
     '画曲线
     curX = int边框左
     curLastY = 0
     For lngLoop = LBound(varLindData) To UBound(varLindData)
         strLineData = varLindData(lngLoop)
         Do While strLineData <> ""
             If InStr(strLineData, ",") > 0 Then
                 curK = Val(Mid(strLineData, 1, InStr(strLineData, ",") - 1))
                 strLineData = Mid(strLineData, InStr(strLineData, ",") + 1)
             Else
                 curK = Val(strLineData)
                 strLineData = ""
             End If
             curLastY = curY
         
             curX = curX + 1
         
             curY = (intHight - curK) + int边框上
             If curX > int边框左 + 1 And curLastY < intHight + int边框上 - 2 * mcurYunti Then objPic.Line (curX, curY)-(curX - 1, curLastY), vbBlue
         Loop
     Next
     objPic.DrawWidth = 1        '线宽
     '画座标
     lngColor = vbBlack
     objPic.Line (int边框左, int边框上 + intHight)-(int边框左 + intWidth, int边框上 + intHight), lngColor
     objPic.Line (int边框左, int边框上 + intHight)-(int边框左, int边框上), lngColor
 
     '刻度
     With objPic
         .FontName = "宋体"
         .ForeColor = lngColor
         .FontBold = False
         .FontSize = 9
     End With
 
     If InStr(str刻度, "|") > 0 Then
         var刻度 = Split(Split(str刻度, "|")(0), ",")
     Else
         var刻度 = Split(str刻度, ",")
     End If
     For lngLoop = LBound(var刻度) To UBound(var刻度)
         curK = Val(Split(var刻度(lngLoop), "-")(0))
         strK = Split(var刻度(lngLoop), "-")(1)
         Call DrawK_X(objPic, int边框左, int边框上 + intHight, curK, strK)
     Next
     If InStr(str刻度, "|") > 0 Then
         var刻度 = Split(Split(str刻度, "|")(1), ",")
         For lngLoop = LBound(var刻度) To UBound(var刻度)
             curK = Val(Split(var刻度(lngLoop), "-")(0))
             strK = Split(var刻度(lngLoop), "-")(1)
             Call DrawK_Y(objPic, int边框左, int边框上 + intHight, curK, strK, 5 / mcurYunti)
         Next
     End If

     '画虚线
     objPic.DrawWidth = 1
     objPic.DrawStyle = vbDot
     Do While str界标 <> ""
         If InStr(str界标, ",") > 0 Then
             curK = Val(Mid(str界标, 1, InStr(str界标, ",") - 1))
             str界标 = Mid(str界标, InStr(str界标, ",") + 1)
         Else
             curK = Val(str界标)
             str界标 = ""
         End If
         If curK <> 0 Then Call DrawK_X(objPic, int边框左, int边框上 + intHight - 10 * mcurYunti, curK, "", Val(intHight - 20 * mcurYunti))
     Loop

     '标题
     If Trim(str标题) <> "" Then
         With objPic
             .CurrentX = int边框左 + intWidth - (Len(str标题) * 12 / Xpixe)
             .CurrentY = int边框上 - int边框上 + 5 * mcurYunti
             .FontSize = 10
             .FontBold = True
         End With
         objPic.Print Trim(str标题)
     End If
     objPic.Scale (0, 0)-(curOldSW, curOldSH)
     objPic.Width = curOldW
     objPic.Height = curOldH
     DrawGam = True
        Exit Function
errHandle:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Sub DrawK_X(ByRef objPic As PictureBox, ByVal curX As Currency, ByVal curY As Currency, ByVal curK As Currency, ByVal strK As String, Optional curHight As Currency = 10)
    '画X轴刻度
    objPic.Line (curX + curK, curY)-(curX + curK, curY - curHight)
    If strK <> "" Then Call PrintRotText(objPic.hDC, strK, (curX + curK) * Xpixe, curY * Ypixe + 8, 0)
End Sub

Private Sub DrawK_Y(ByRef objPic As PictureBox, ByVal curX As Currency, ByVal curY As Currency, ByVal curK As Currency, ByVal strK As String, Optional curWidth As Currency = 10)
    '画Y轴刻度
    objPic.Line (curX, curY - curK)-(curX + curWidth, curY - curK)
    If strK <> "" Then Call PrintRotText(objPic.hDC, strK, curX * Xpixe - 11, (curY - curK) * Ypixe, 0)
End Sub




