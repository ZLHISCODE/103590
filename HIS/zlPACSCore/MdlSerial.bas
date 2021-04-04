Attribute VB_Name = "MdlSerial"
Option Explicit

'--------------------------------------------------------
'功  能：本模块为序列相关内容处理
'编制日期：2004.6
'过程函数清单：
'funSliceLocation   ():在一个Viewer中寻找切片位置最近指定值的图像
'subSerialPlaceInPhase  ():序列之间位置同步
'FunImageIsX        ():判断一个图像所在的行
'FunImageIsY        ():判断一个图像所在的列
'subIsSerialXY      ():判断当前点在哪个序列的位置上
'subDispframe       ():显示当前窗体指定viewer的图像外框，选择标记等
'subInitSerial      ():删除原有拖动条，并重新添加横向，纵向，双向拖动条。'
'
'修改记录：
'    2005.6     黄捷
'-------------------------------------------------------

Private Function funSliceLocation(intViewerIndex As Integer, s As Double) As Integer
'------------------------------------------------
'功能：在一个Viewer中寻找切片位置最近指定值的图像
'参数：v--寻找最近切片图像的viewer；s--进行比较的目标切片位置。
'返回：最近的图像序号。
'------------------------------------------------
    Dim dt As Double
    Dim i As Integer
    
    dt = intSliceOffset
    funSliceLocation = 0
    For i = 1 To ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
        '查找最接近s的位置
        If Abs(Val(ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).SliceLocation) - s) < dt Then
            dt = Abs(Val(ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).SliceLocation) - s)
            funSliceLocation = i
        End If
    Next
End Function

Public Sub subManualSeriesSyn(f As frmViewer, iMove As Integer, vIndex As Integer)
'------------------------------------------------
'功能：手工序列间位置同步
'参数： f--进行序列同步的窗体
'       iMove--图像翻动的方向和数量，正数向前翻动，负数向后翻动。
'       vIndex 是哪个viewer启动手工序列同步
'返回：无，直接修改图像的显示。
'------------------------------------------------
    Dim v As DicomViewer
    Dim i As Integer
    Dim intOldIndex As Integer
    Dim intCurrentIndex As Integer
    
    f.blnVscroInvoked = True
    For Each v In f.Viewer
        If v.Visible And v.Index <> vIndex Then
            If ZLShowSeriesInfos(v.Index).Selected = True Then
                '保存现在的图像位置
                intOldIndex = f.VScro(v.Index).Value
                
                '计算同步后图像的新位置
                If (intOldIndex + iMove) <= (ZLShowSeriesInfos(v.Index).ImageInfos.Count - v.MultiColumns * v.MultiRows + 1) And (intOldIndex + iMove) > 0 Then
                    intCurrentIndex = intOldIndex + iMove
                ElseIf iMove > 0 And intOldIndex < ZLShowSeriesInfos(v.Index).ImageInfos.Count Then
                    intCurrentIndex = ZLShowSeriesInfos(v.Index).ImageInfos.Count - v.MultiColumns * v.MultiRows + 1
                ElseIf iMove < 0 And intOldIndex > 1 Then
                    intCurrentIndex = 1
                ElseIf iMove > 0 And intOldIndex = ZLShowSeriesInfos(v.Index).ImageInfos.Count Then
                    intCurrentIndex = ZLShowSeriesInfos(v.Index).ImageInfos.Count
                End If
                
                If intCurrentIndex <= 0 Or intCurrentIndex > ZLShowSeriesInfos(v.Index).ImageInfos.Count Then
                    intCurrentIndex = 1
                End If
                
                If intOldIndex <> intCurrentIndex Then       '图像发生翻动
                    f.MSFViewer.TextMatrix(v.Index, 3) = intCurrentIndex
                    f.VScro(v.Index).Value = intCurrentIndex
                    Call subShowALLImage(f, v, intCurrentIndex, True)
                End If
            End If
        End If
        
    Next
    f.blnVscroInvoked = False
End Sub

Public Sub subSerialPlaceInPhase(dubPlace As Double, f As frmViewer)
'------------------------------------------------
'功能：序列之间位置同步
'参数：dubPlace--进行序列同步的目的切片位置；f--进行序列之间位置同步的窗体。
'返回：无，直接调整图像位置。
'------------------------------------------------
    Dim v As DicomViewer, i As Integer
    Dim m As Integer
    Dim intCurrentIndex As Integer
    
    f.blnVscroInvoked = True
    For Each v In f.Viewer
        If v.Index <> f.intSelectedSerial And v.Visible Then
            If ZLShowSeriesInfos(v.Index).Selected = True Then
            intCurrentIndex = funSliceLocation(v.Index, dubPlace)
                If intCurrentIndex <> 0 Then
                    m = intCurrentIndex - f.MSFViewer.TextMatrix(v.Index, 3)
                    f.MSFViewer.TextMatrix(v.Index, 3) = intCurrentIndex
                    '有滚动条的情况
                    If ZLShowSeriesInfos(v.Index).ImageInfos.Count > v.MultiColumns * v.MultiRows Then
                        If intCurrentIndex > ZLShowSeriesInfos(v.Index).ImageInfos.Count - v.MultiColumns * v.MultiRows + 1 Then
                            intCurrentIndex = ZLShowSeriesInfos(v.Index).ImageInfos.Count - v.MultiColumns * v.MultiRows + 1
                        End If
                        If intCurrentIndex < 1 Then intCurrentIndex = 1
                        
                        '图像发生翻动
                        If f.VScro(v.Index).Value <> intCurrentIndex Then
                            f.MSFViewer.TextMatrix(v.Index, 3) = intCurrentIndex
                            f.VScro(v.Index).Value = intCurrentIndex
                            Call subShowALLImage(f, v, intCurrentIndex, True)
                        End If
                    End If
                End If
            End If
        End If
    Next
    f.blnVscroInvoked = False
End Sub

Function FunImageIsX(Index As Integer, v As DicomViewer) As Integer
'''''判断一个图像所在的行
'2009用
    FunImageIsX = Index - v.CurrentIndex + 1
    FunImageIsX = FunImageIsX Mod v.MultiColumns
    If FunImageIsX = 0 Then FunImageIsX = v.MultiColumns
End Function

Function FunImageIsY(Index As Integer, v As DicomViewer) As Integer
'''''判断一个图像所在的列
'2009用
    FunImageIsY = Index - v.CurrentIndex + 1
    FunImageIsY = Int(FunImageIsY / v.MultiColumns - 0.5) + 1
End Function


Public Sub subIsSerialXY(f As frmViewer, x, y, intSerialX As Integer, intSerialY As Integer)
'------------------------------------------------
'功能：判断当前点在哪个序列的位置上
'参数：f--进行判断的窗体；(x,y)--需要判断的点的坐标；
'      intSerialX--返回该点所在x方向的序列数；intSerialY--返回该点所在y方向的序列数。
'返回：无
'2009用
'------------------------------------------------
    Dim i As Integer, j As Integer
    Dim tx1 As Integer, tx2 As Integer, ty1 As Integer, ty2 As Integer
    intSerialX = 0
    intSerialY = 0
    For i = 1 To intMaxAreaX                          ''横向最多可划分的区域
        If i = 1 Then
            tx1 = 0
        Else
            tx1 = f.PicX(i - 1).left
        End If
        If i = intMaxAreaX Then
            tx2 = f.picViewer.ScaleWidth
        Else
            tx2 = f.PicX(i).left
        End If
        For j = 1 To intMaxAreaY
            If j = 1 Then
                ty1 = 0
            Else
                ty2 = f.PicY(j - 1).top
            End If
            If j = intMaxAreaY Then
                ty2 = f.picViewer.ScaleHeight
            Else
                ty2 = f.PicY(j).top
            End If
            If x >= tx1 And x <= tx2 And y >= ty1 And y <= ty2 Then
                intSerialX = i
                intSerialY = j
                Exit Sub
            End If
        Next
    Next
End Sub

Public Sub subDispframe(f As frmViewer, v As DicomViewer)
'------------------------------------------------
'功能：显示当前窗体指定viewer的图像外框，选择标记等
'参数：f--需要显示viewer图像外框的窗体；v--需要显示图像外框的viewer。
'返回：无，直接显示指定viewer的图像外框。
'2009用
'------------------------------------------------
    Dim l, lx, ly, lb As DicomLabel
    Dim x, y As Integer
    Dim w, h As Single
    Dim i As Integer
    Dim iTempIndex As Integer
    
    If v.Index = 0 Then Exit Sub
    If v.Images.Count = 0 Then Exit Sub
    
    On Error GoTo err
    
    '清除原有的标注
    v.Labels.Clear
    
    '计算每一个图像区域的宽度和高度
    w = v.width / v.MultiColumns / Screen.TwipsPerPixelX
    h = v.height / v.MultiRows / Screen.TwipsPerPixelY
    
    '显示打印标记
    For i = v.CurrentIndex To v.CurrentIndex + v.MultiColumns * v.MultiRows
        If i > v.Images.Count Then Exit For
        v.Images(i).Labels(G_INT_SYS_LABEL_PRINT_TAG).Visible = (blnShowPrintTag And ZLShowSeriesInfos(v.Index).ImageInfos(v.Images(i).Tag).blnPrinted)
    Next i
    
    '循环每一个图像，画标注
    For y = 1 To v.MultiRows
        For x = 1 To v.MultiColumns
            '计算当前的图像Index
            iTempIndex = x + (y - 1) * v.MultiColumns + v.CurrentIndex - 1
            ''''''''''''''''边框''''''''''''''
            Set l = New DicomLabel
            l.LabelType = 2         '矩形框
            l.width = w - lngCellSpacing * 2
            l.height = h - lngCellSpacing * 2
            l.left = (x - 1) * w + lngCellSpacing
            l.top = (y - 1) * h + lngCellSpacing
            l.Tag = "L" & x + (y - 1)
            '''判断是否当前选择的图像，来决定图像矩形框的颜色和线型，线宽
            '当前图像边框颜色 lngCurrentImageBorderColor
            '选中图像边框颜色 lngSelectedImageBorderColor
            '当前（未选中）序列边框颜色 lngCurrentSeriesBorderColor
            '先判断这个序列是否被选中
            If ZLShowSeriesInfos(v.Index).Selected = True Then   '被选中的序列
                '在判断当前图像是否是当前图像
                If v.Index = f.intSelectedSerial And iTempIndex = f.MSFViewer.TextMatrix(v.Index, 3) Then
                    l.ForeColour = lngCurrentImageBorderColor
                    l.LineStyle = lngCurrentImageBorderLineStyle
                    l.LineWidth = lngCurrentImageBorderLineWidth
                Else
                    l.ForeColour = lngSelectedImageBorderColor
                    l.LineStyle = lngSelectedImageBorderLineStyle
                    l.LineWidth = lngSelectedImageBorderLineWidth
                End If
            ElseIf v.Index = f.intSelectedSerial Then   '没有被选中，则判断是否是当前序列
                '在判断当前图像是否是当前图像
                If iTempIndex = f.MSFViewer.TextMatrix(v.Index, 3) Then
                    l.ForeColour = lngCurrentImageBorderColor
                Else
                    l.ForeColour = lngCurrentSeriesBorderColor
                End If
                l.LineStyle = lngCurrentImageBorderLineStyle
                l.LineWidth = lngCurrentImageBorderLineWidth
            Else        '既没有被选中，也不是当前序列，则显示默认边框
                l.ForeColour = lngDefaultImageBorderColor
                l.LineStyle = lngDefaultImageBorderLineStyle
                l.LineWidth = lngDefaultImageBorderLineWidth
            End If
            v.Labels.Add l
            
            '如果是当前序列的当前图像，则为这个图像中被选中的标注显示8个标注选择句柄
            If iTempIndex <= v.Images.Count And iTempIndex > 0 Then
                If v.Images(iTempIndex).Labels(11).Visible Then
                    '为指定图像中的指定标注，显示标注选择句柄
                    SubDispPeriod v.Images(iTempIndex).Labels(11).TagObject, v.Images(iTempIndex), f
                End If
            End If
            
            ''''''''''''''''为每一个图像增加选择标记：短横线；短竖线；选择标记''''''''''''''
            '''''''''''''''''''横线'''''''''''''''''''''''''''''''''
            Set lx = New DicomLabel
            lx.LabelType = 3            '直线
            lx.width = lngImageIdentifierSize
            lx.height = 0
            lx.left = l.left + l.width - lngImageIdentifierSize
            lx.top = l.top + l.height - lngImageIdentifierSize
            lx.Tag = "X" & x + (y - 1) * v.MultiColumns
            lx.TagObject = l
            lx.ForeColour = l.ForeColour
            lx.LineStyle = l.LineStyle
            lx.LineWidth = l.LineWidth
            v.Labels.Add lx
            ''''''''''''''''竖线''''''''''''''
            Set ly = New DicomLabel
            ly.LabelType = 3        '直线
            ly.width = 0
            ly.height = lngImageIdentifierSize
            ly.left = l.left + l.width - lngImageIdentifierSize
            ly.top = l.top + l.height - lngImageIdentifierSize
            ly.Tag = "Y" & x + (y - 1) * v.MultiColumns
            ly.TagObject = lx
            ly.ForeColour = l.ForeColour
            ly.LineStyle = l.LineStyle
            ly.LineWidth = l.LineWidth
            v.Labels.Add ly
            ''''''''''''''''选择标记''''''''''''''
            Set lb = New DicomLabel
            lb.LabelType = 2            '矩形
            lb.width = lngImageIdentifierSize - IIf(l.LineWidth / 2 >= 2, l.LineWidth / 2, 2)
            lb.height = lngImageIdentifierSize - IIf(l.LineWidth / 2 >= 2, l.LineWidth / 2, 2)
            lb.left = l.left + l.width - lngImageIdentifierSize + 1
            lb.top = l.top + l.height - lngImageIdentifierSize + 1
            lb.Transparent = False
            lb.ForeColour = lngSelectImageForeColour
            lb.BackColour = lngSelectImageForeColour
            lb.Tag = "B" & x + (y - 1) * v.MultiColumns
            lb.TagObject = ly
            lb.Visible = False
            v.Labels.Add lb
            l.TagObject = lb
            If (iTempIndex <= v.Images.Count) Then
                If v.Images.Count > 0 And (iTempIndex) > 0 Then
                    lb.Visible = ZLShowSeriesInfos(v.Index).ImageInfos(v.Images(iTempIndex).Tag).blnSelected
                    If iTempIndex <> v.Images(iTempIndex).Tag Then
                        Debug.Print "ii"
                    End If
                Else
                    lb.Visible = False
                End If
            End If
            If (iTempIndex = v.Images.Count) And Not blnDsipSpilthBorder Then Exit Sub
        Next
    Next
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subDispWWWL(img As DicomImage)
'------------------------------------------------
'功能：按照预设的窗宽窗位值显示图像
'参数：img--需要重新显示的图像
'返回：无，直接修改图像的窗宽窗位值，和窗宽窗位标注
'上级函数或过程：
'下级函数或过程：无
'引用的外部参数：aPresetWinWL
'编制人：黄捷
'------------------------------------------------
    Dim strDriverType As String
    Dim intModality As Integer
    Dim i As Integer
    Dim im As DicomImage
    
    If IsNull(img.Attributes(&H8, &H60).Value) Then Exit Sub         '获取Modality
    strDriverType = img.Attributes(&H8, &H60).Value
    
    For i = 1 To UBound(aPresetWinWL, 2)
        If UCase(aPresetWinWL(3, i).strModality) = UCase(strDriverType) Then
            intModality = i
            Exit For
        End If
    Next i
    
    For i = 3 To 12
        If aPresetWinWL(i, intModality).bInUse And aPresetWinWL(i, intModality).intDefault = 1 Then
            img.width = aPresetWinWL(i, intModality).lngWinWidth
            img.Level = aPresetWinWL(i, intModality).lngWinLevel
            Exit For
        End If
    Next i
    img.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & img.width & "-L:" & img.Level
End Sub

Public Sub subScaleImage(img As DicomImage, v As DicomViewer, lngOldX As Long, lngOldY As Long)
'------------------------------------------------
'功能：修正图像的位置和缩放比例，在图像布局有所改变的时候调用
'参数： img         --- 需要调整的图像
'       v           --- 图像所在的新摆放好的Viewer
'       lngOldX     --- 图像原来所在的Viewer中单个图像所占用的宽度
'       lngOldY     --- 图像原来所在的Viewer中单个图像所占用的高度
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim dblScale As Double
    Dim dblNewZoom As Double
    Dim lngNewX As Long
    Dim lngNewY As Long
    Dim dblScaleX As Double
    Dim dblScaleY As Double
    Dim dblOldXY As Double
    Dim dblNewXY As Double
    
    If lngOldX = 0 Or lngOldY = 0 Then Exit Sub
    
    '计算新Viewer中单个图像所占用的宽度和高度
    lngNewX = v.width / v.MultiColumns
    lngNewY = v.height / v.MultiRows
    
    '计算缩放的比例
'    dblScale = lngNewX / lngOldX
'    If Abs(dblScale - 1) > Abs(lngNewY / lngOldY - 1) Then
'        dblScale = lngNewY / lngOldY
'    End If
    '计算缩放比例
    '如果Sx和Sy都大于1，则取小的。
    '如果Sx和Sy都小于1，取大的。
    '如果Sx大于1，Sy小于1或者Sx小于1，Sy大于1，取OX/OY，NX/NY中比例大的短边为标准
    dblScaleX = lngNewX / lngOldX
    dblScaleY = lngNewY / lngOldY
    If dblScaleX > 1 And dblScaleY > 1 Then
        dblScale = IIf(dblScaleX < dblScaleY, dblScaleX, dblScaleY)
    ElseIf dblScaleX < 1 And dblScaleY < 1 Then
        dblScale = IIf(dblScaleX > dblScaleY, dblScaleY, dblScaleX)
    Else
        If lngOldX >= lngOldY Then
            dblOldXY = lngOldX / lngOldY
        Else
            dblOldXY = lngOldY / lngOldX
        End If
        If lngNewX >= lngNewY Then
            dblNewXY = lngNewX / lngNewY
        Else
            dblNewXY = lngNewY / lngNewX
        End If
        
        If dblOldXY >= dblNewXY Then
            dblScale = IIf(lngOldX < lngOldY, dblScaleX, dblScaleY)
        Else
            dblScale = IIf(lngNewX < lngNewY, dblScaleX, dblScaleY)
        End If
    End If
    
    '计算新的Zoom
    dblNewZoom = dblScale * img.ActualZoom
    
    '先调整Scroll
    img.ScrollX = img.ActualScrollX - (lngNewX - lngOldX) / 2 / Screen.TwipsPerPixelX
    img.ScrollY = img.ActualScrollY - (lngNewY - lngOldY) / 2 / Screen.TwipsPerPixelY
    
    '用居中缩放的方式调整Zoom
    Call subCenterZoom(img, v, dblNewZoom)
End Sub

Public Sub subInitSerial(f As frmViewer)
'------------------------------------------------
'功能：unload原有的分隔条，然后重新load分隔条，并把它们摆放到初始位置
'参数：f--需要初始化拖动条的窗体
'返回：无，直接对窗体中的拖动条进行初始化。
'2009用
'------------------------------------------------
    Dim i, j, k As Long
    With f
        '清除原有的三种拖动：横向，纵向，双向。
        For i = 1 To .PicX.Count - 1
            Unload .PicX(i)
        Next
        For i = 1 To .PicY.Count - 1
            Unload .PicY(i)
        Next
        For i = 1 To .PicXY.Count - 1
            Unload .PicXY(i)
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        .PicX(0).width = intSpaceSize
        .PicX(0).height = .picViewer.height
        .PicX(0).top = 0
        .PicXX.width = intSpaceSize
        .PicXX.height = .picViewer.height
        .PicXX.top = 0
        .PicY(0).height = intSpaceSize
        .PicY(0).width = .picViewer.width
        .PicY(0).left = 0
        .PicYY.height = intSpaceSize
        .PicYY.width = .picViewer.width
        .PicYY.left = 0
        .PicXY(0).height = intSpaceSize
        .PicXY(0).width = intSpaceSize
        .PicXY(0).top = .PicY(0).top
        .PicXY(0).left = .PicX(0).left
        .PicY(0).AutoRedraw = True
        .PicX(0).AutoRedraw = True
        .PicXY(0).AutoRedraw = True
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 1 To intMaxAreaX - 1 ''''初始化横向间隔
            load .PicX(i)
            .PicX(i).left = .picViewer.width - intSpaceSize
            .PicX(i).AutoRedraw = True
            Call zlControl.PicShowFlat(.PicX(i), 1)     '将PictureBox模拟成3D平面按钮
            .PicX(i).Visible = True
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 1 To intMaxAreaY - 1 ''''初始化纵向间隔
            load .PicY(i)
            .PicY(i).top = .picViewer.height - intSpaceSize
            .PicY(i).AutoRedraw = True
             Call zlControl.PicShowFlat(.PicY(i), 1)        '将PictureBox模拟成3D平面按钮
            .PicY(i).Visible = True
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = intMaxAreaX - 1 To 1 Step -1 ''''初始化横纵向间隔
            For j = intMaxAreaY - 1 To 1 Step -1
                k = (j - 1) * (intMaxAreaX - 1) + i
                load .PicXY(k)
                .PicXY(k).top = .PicY(j).top
                .PicXY(k).left = .PicX(i).left
                .PicXY(k).Visible = True
                .PicXY(k).ZOrder
            Next
        Next
    End With
End Sub

