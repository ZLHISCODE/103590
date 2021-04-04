Attribute VB_Name = "mdlLabel"
Option Explicit
'--------------------------------------------------------
'功  能：本模块为标注处理的函数过程等
'编制人：胡涛
'编制日期：2005.06.12
'过程函数清单：
'    funAccountMiddlePoint():   直线上两个点和第三个点的X或Y坐标，计算出对应的Y或X坐标。
'    subPeriodMovee5X():        矢冠状中心点的移动
'    subPeriodMove():           矢冠状控制点中四个边点的移动处理
'    funROIResultString():      生成测量结果字符串,根据系统配置中设置的是否显示面积，平均值，均方差等条件
'    subMove25():               裁减根据系统标注1，移动标注2-5。
'    subTakeOut1():             从选中的标注中去掉系统标注用于裁减的序号为2-5的标注，对于序号1的标注是否删除，通过参数isTakeOut1来判断
'    funLabelType():            返回内部使用的标注类型名称
'    SubDispPeriod():           为指定图像中的指定标注，显示标注选择句柄
'    SubDispLinePeriod():       为图像中的指定直线和箭头标注，显示标注选择句柄
'    subCutOutInphase():        在裁减状态下对裁减操作作出图像同步处理
'    funMouseOverPeriod():      返回鼠标所越过的句柄编号
'    subMoveMPRLabel():         移动矢冠状重建控制点、线，且生成新的重建图像。
'    subMoveLable():            移动一个标注,包括矢冠状重建标注、用户标注和裁剪标注
'    subChangeLableSize():      改变一个标注的大小,并修改其相关测量信息的显示值
'    SubNoDispPeriod():         为指定图像隐藏标注选择句柄
'    subTextCoordinate():       根据图像的反转情况决定文字的坐标转换
'    SubChangeColor():          改变选中LABEL的颜色
'    GetNewLabel():             生成一个LABEL对象，并对其做初始化。
'    subDeleteAppointLabel():   删除指定类型的标注
'    SubInitPeriod():           为每一幅图增加前n个系统句柄，n的数量由常量G_INT_SYS_LABEL_COUNT决定
'    UpdateMarkers():           根据图像显示或隐藏病人体位信息
'    UpdateRuler():             显示图像标尺
'    subDispImageInfo():        显示或隐藏病人图像四角信息和窗宽窗位显示
'    subGetImgInfoLabel():      从图像中提取病人的四个角信息标注，配合系统参数设置中四个角标注的内容使用
'    subInitImageLabels():      初始化、显示或隐藏指定图像的标注信息:系统标注；体位标注，标尺，四角信息，窗宽窗位
'    funcCalImgInfoLabel():     根据传入的中文简称，来计算对应四角标注的显示值。
'    subSaveLabelToImg():       将标注保存到DICOM图像的头信息里面
'    subReadLabelFromImg():     从图像的头文件中读取标注，并显示标注
'    funDrawVas():              根据lblLine做自动血管测量
'修改记录：
'    2005.06.30    黄捷      程序优化
'-------------------------------------------------------

Private Function funAccountMiddlePoint(x1, y1, x2, y2, X3)
'------------------------------------------------
'功能：直线上两个点和第三个点的X或Y坐标，计算出对应的Y或X坐标。
'参数：(X1,Y1)--直线上第一个点的坐标；（X2，Y2）--直线上第二个点的坐标；X3--直线上第三个点的X坐标
'返回：第三个点的Y或X坐标
'2009用
'------------------------------------------------
    funAccountMiddlePoint = 0
    If x1 = x2 Then Exit Function
    If y1 = y2 Then
        funAccountMiddlePoint = y1
        Exit Function
    End If
    funAccountMiddlePoint = (X3 - x2) * (y2 - y1) / (x2 - x1) + y2
    If funAccountMiddlePoint > 30000 Then funAccountMiddlePoint = 30000
    If funAccountMiddlePoint < -30000 Then funAccountMiddlePoint = -30000
End Function

Public Sub subPeriodMovee5X(la, lb, ll, xx, Yy, E5, basex, baseY, im As DicomImage)
'------------------------------------------------
'功能：矢冠状中心点的移动
'参数：la--与中心控制点在同一直线上的边控制点；lb--与边控制点la在同一直线的另一个边控制点；
'      ll--la与lb所连接的控制线；xx--中心控制点中心点的新X位置；Yy--中心控制点中心点的新Y位置；
'      E5--中心控制点；basex--中心控制点中心点的旧X位置；baseY--中心控制点中心点的旧Y位置；
'      im--包含上述控制点线的图像
'返回：无。直接移动了la，lb，ll三个控制点和线
'2009用
'------------------------------------------------
    Dim x, y
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If ll.height = 0 Then
        la.top = E5.top
        lb.top = E5.top
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf ll.width = 0 Then
        la.left = E5.left
        lb.left = E5.left
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf la.top = -G_INT_MPR_RADIUS / 2 Then   '''''如果la在顶上
        x = funAccountMiddlePoint(la.top + Yy - baseY, la.left, E5.top, E5.left, -G_INT_MPR_RADIUS / 2)
        x = x + xx - basex
        la.left = x
        If la.left + G_INT_MPR_RADIUS / 2 < 0 Then   '''la超出左边
            la.top = funAccountMiddlePoint(la.left, -G_INT_MPR_RADIUS / 2, E5.left, E5.top, -G_INT_MPR_RADIUS / 2)
            la.left = -G_INT_MPR_RADIUS / 2
        ElseIf la.left + G_INT_MPR_RADIUS / 2 > im.sizex Then   '''la超出右边
            la.top = funAccountMiddlePoint(la.left, -G_INT_MPR_RADIUS / 2, E5.left, E5.top, im.sizex - G_INT_MPR_RADIUS / 2)
            la.left = im.sizex - G_INT_MPR_RADIUS / 2
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf la.top = im.sizey - G_INT_MPR_RADIUS / 2 Then  ''如果la在底下
        x = funAccountMiddlePoint(la.top + Yy - baseY, la.left, E5.top, E5.left, im.sizey - G_INT_MPR_RADIUS / 2)
        x = x + xx - basex
        la.left = x
        If la.left + G_INT_MPR_RADIUS / 2 < 0 Then
            la.top = funAccountMiddlePoint(la.left, im.sizey - G_INT_MPR_RADIUS / 2, E5.left, E5.top, -G_INT_MPR_RADIUS / 2)
            la.left = -G_INT_MPR_RADIUS / 2
        ElseIf la.left + G_INT_MPR_RADIUS / 2 > im.sizex Then
            la.top = funAccountMiddlePoint(la.left, im.sizey - G_INT_MPR_RADIUS / 2, E5.left, E5.top, im.sizex - G_INT_MPR_RADIUS / 2)
            la.left = im.sizex - G_INT_MPR_RADIUS / 2
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf la.left = -G_INT_MPR_RADIUS / 2 Then    ''如果la在左边
        y = funAccountMiddlePoint(la.left + xx - basex, la.top, E5.left, E5.top, -G_INT_MPR_RADIUS / 2)
        y = y + Yy - baseY
        la.top = y
        If la.top + G_INT_MPR_RADIUS / 2 < 0 Then   '''la超出上边
            la.left = funAccountMiddlePoint(la.top, -G_INT_MPR_RADIUS / 2, E5.top, E5.left, -G_INT_MPR_RADIUS / 2)
            la.top = -G_INT_MPR_RADIUS / 2
        ElseIf la.top + G_INT_MPR_RADIUS / 2 > im.sizex Then  '''la超出下边
            la.left = funAccountMiddlePoint(im.sizex - G_INT_MPR_RADIUS / 2, la.left, E5.top, E5.left, -G_INT_MPR_RADIUS / 2)
            la.top = im.sizex - G_INT_MPR_RADIUS / 2
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf la.left = im.sizey - G_INT_MPR_RADIUS / 2 Then   ''如果la在右边
        y = funAccountMiddlePoint(la.left + xx - basex, la.top, E5.left, E5.top, im.sizey - G_INT_MPR_RADIUS / 2)
        y = y + Yy - baseY
        la.top = y
        If la.top + G_INT_MPR_RADIUS / 2 < 0 Then   '''la超出上边
            la.left = funAccountMiddlePoint(la.top, -G_INT_MPR_RADIUS / 2, E5.top, E5.left, im.sizex - G_INT_MPR_RADIUS / 2)
            la.top = -G_INT_MPR_RADIUS / 2
        ElseIf la.top + G_INT_MPR_RADIUS / 2 > im.sizex Then  '''la超出下边
            la.left = funAccountMiddlePoint(im.sizex - G_INT_MPR_RADIUS / 2, la.left, E5.top, E5.left, im.sizex - G_INT_MPR_RADIUS / 2)
            la.top = im.sizex - G_INT_MPR_RADIUS / 2
        End If
    End If
    ''''''计算对点'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If la.top < E5.top Then
        lb.left = funAccountMiddlePoint(la.top, la.left, E5.top, E5.left, im.sizey - G_INT_MPR_RADIUS / 2)
        lb.top = im.sizey - G_INT_MPR_RADIUS / 2
        If lb.left + G_INT_MPR_RADIUS / 2 < 0 Then
            lb.top = funAccountMiddlePoint(lb.left, lb.top, E5.left, E5.top, G_INT_MPR_RADIUS / 2)
            lb.left = -G_INT_MPR_RADIUS / 2
        ElseIf lb.left + G_INT_MPR_RADIUS / 2 > im.sizex Then
            lb.top = funAccountMiddlePoint(lb.left, lb.top, E5.left, E5.top, im.sizex - G_INT_MPR_RADIUS / 2)
            lb.left = im.sizex - G_INT_MPR_RADIUS / 2
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf la.top > E5.top Then
        lb.left = funAccountMiddlePoint(la.top, la.left, E5.top, E5.left, -G_INT_MPR_RADIUS / 2)
        lb.top = -G_INT_MPR_RADIUS / 2
        If lb.left + G_INT_MPR_RADIUS / 2 < 0 Then
            lb.top = funAccountMiddlePoint(lb.left, lb.top, E5.left, E5.top, -G_INT_MPR_RADIUS / 2)
            lb.left = -G_INT_MPR_RADIUS / 2
        ElseIf lb.left + G_INT_MPR_RADIUS / 2 > im.sizex Then
            lb.top = funAccountMiddlePoint(lb.left, lb.top, E5.left, E5.top, im.sizex - G_INT_MPR_RADIUS / 2)
            lb.left = im.sizex - G_INT_MPR_RADIUS / 2 - G_INT_MPR_RADIUS / 2
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ll.left = la.left + G_INT_MPR_RADIUS / 2
    ll.top = la.top + G_INT_MPR_RADIUS / 2
    ll.width = lb.left - la.left
    ll.height = lb.top - la.top
End Sub


Public Sub subPeriodMove(ByVal la As DicomLabel, ByVal x As Long, ByVal y As Long, ByVal lb As DicomLabel, _
                  ByVal ll As DicomLabel, E5 As DicomLabel, im As DicomImage)
'------------------------------------------------
'功能：矢冠状控制点中四个边点的移动处理
'参数： la--被移动的控制边点标注 ；x--标注新位置在图像上的X坐标  ；y--标注新位置在图像上的Y坐标  ；
'       lb--跟被移动标注在同一直线上的另一个控制边点；ll--跟la相连的矢冠状控制线；
'       E5--矢冠状控制点中的中心点；im--做矢冠状重建的图像
'返回：无，直接移动标注，改变了la,lb,ll的位置
'2009用
'------------------------------------------------
    Dim x1 As Long, y1 As Long
    Dim x2 As Long, y2 As Long  '（X2,Y2）记录矢冠状重建中心控制点标注的中心坐标
    Dim X3 As Long, Y3 As Long, movex As Long
    x2 = E5.left + G_INT_MPR_RADIUS / 2    '''中心点位置
    y2 = E5.top + G_INT_MPR_RADIUS / 2
    '''''''''''''''''计算点位置'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If x = x2 Then             ''''鼠标X位置和中心点平行
        X3 = x
        Y3 = IIf(y > y2, 0, 0 + im.sizey)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf y = y2 Then           ''''鼠标Y位置和中心点平行
        Y3 = y
        X3 = IIf(x > x2, 0, 0 + im.sizex)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf y > y2 Then              ''''鼠标位于中心点上方
        Y3 = 0 + im.sizey
        X3 = (Y3 - y2) * (x2 - x) / (y2 - y) + x2
        If X3 < 0 Then
            Y3 = (y2 - Y3) * (0 - X3) / (x2 - X3) + Y3
            X3 = 0
        ElseIf X3 > 0 + im.sizex Then
            Y3 = (y2 - Y3) * (0 + im.sizex - X3) / (x2 - X3) + Y3
            X3 = 0 + im.sizex
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf y < y2 Then              '''''鼠标位于中心点下方
        Y3 = 0
        X3 = (Y3 - y2) * (x2 - x) / (y2 - y) + x2
        If X3 < 0 Then
            Y3 = (y2 - Y3) * (0 - X3) / (x2 - X3) + Y3
            X3 = 0
        ElseIf X3 > 0 + im.sizex Then
            Y3 = (y2 - Y3) * (0 + im.sizex - X3) / (x2 - X3) + Y3
            X3 = 0 + im.sizex
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''计算出第一个标尺的位置
    la.left = X3 - G_INT_MPR_RADIUS / 2
    la.top = Y3 - G_INT_MPR_RADIUS / 2
    '''''''''''''''''''''''''''''''''''''''''''''''''''计算对点位置
    x1 = X3
    y1 = Y3
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If x1 = x2 Then
        X3 = x1
        Y3 = IIf(y1 > y2, 0, 0 + im.sizey)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf y1 = y2 Then
        Y3 = y1
        X3 = IIf(x1 > x2, 0, 0 + im.sizex)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf y1 < y2 Then
        Y3 = 0 + im.sizey
        X3 = (Y3 - y2) * (x2 - x1) / (y2 - y1) + x2
        If X3 < 0 Then
            Y3 = (y2 - Y3) * (0 - X3) / (x2 - X3) + Y3
            X3 = 0
        ElseIf X3 > 0 + im.sizex Then
            Y3 = (y2 - Y3) * (0 + im.sizex - X3) / (x2 - X3) + Y3
            X3 = 0 + im.sizex
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf y1 > y2 Then
        Y3 = 0
        X3 = (Y3 - y2) * (x2 - x1) / (y2 - y1) + x2
        If X3 < 0 Then
            Y3 = (y2 - Y3) * (0 - X3) / (x2 - X3) + Y3
            X3 = 0
        ElseIf X3 > 0 + im.sizex Then
            Y3 = (y2 - Y3) * (0 + im.sizex - X3) / (x2 - X3) + Y3
            X3 = 0 + im.sizex
        End If

    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    lb.left = X3 - G_INT_MPR_RADIUS / 2
    lb.top = Y3 - G_INT_MPR_RADIUS / 2
'    '''''''''''''''''''''''''''''''''''''''''''''
    ll.left = la.left + G_INT_MPR_RADIUS / 2
    ll.top = la.top + G_INT_MPR_RADIUS / 2
    ll.width = X3 - ll.left
    ll.height = Y3 - ll.top
End Sub

Public Function funROIResultString(la As DicomLabel, img As DicomImage) As String
'------------------------------------------------
'功能：生成测量结果中文或英文字符串,根据系统配置中设置的是否显示面积，平均值，均方差等条件，为输入的标注返回其测量结果
'参数：la--为需要测量结果的标注，程序内部根据不同的标注类型返回不同的结果，对于直线，多边线测量，返回测量的长度
'返回：为测量结果字符串
'2009用
'------------------------------------------------
    funROIResultString = ""
    Dim strROIArea As String
    Dim strROIMean As String
    Dim strROIStdDev As String
    Dim strROILength As String
    Dim strROIMax As String
    Dim strROIMin As String
    Dim strAngle As String
    Dim lTemp As DicomLabel
    If bROITextChinese Then        ''使用中文表示测量信息
        strROIArea = "面积："
        strROIMean = "平均值："
        strROIMax = "最大值："
        strROIMin = "最小值："
        If img.Attributes(&H8, &H60).Exists And Not IsNull(img.Attributes(&H8, &H60).Value) Then
            If UCase(img.Attributes(&H8, &H60).Value) = "CT" Then
                strROIMean = "平均CT值："
                strROIMax = "最大CT值："
                strROIMin = "最小CT值："
            End If
        End If
        
        strROIStdDev = "标准差："
        strROILength = "周长："
        strAngle = "角度："
    Else
        strROIArea = "Area: "
        strROIMean = "Mean: "
        strROIStdDev = "Std.Dev: "
        strROILength = "Length:"
        strAngle = "Angle:"
        strROIMax = "Max: "
        strROIMin = "Min:"
    End If
    
    '屏蔽错误,主要是防止伪彩图象出现错误
    On Error Resume Next
    
    If left(la.Tag, 2) = "JD" Then
        '计算角度
        If bROIArea Then
            Set lTemp = la
            If lTemp.Tag = "JD1" Then
                If Not lTemp.TagObject.TagObject Is Nothing Then
                    Set lTemp = lTemp.TagObject.TagObject
                End If
            End If
            funROIResultString = strAngle & Int(GetAngle(lTemp.left, lTemp.top, lTemp.left + lTemp.width, lTemp.top + lTemp.height, lTemp.TagObject.left, lTemp.TagObject.top) * 100) / 100
        End If
    ElseIf la.LabelType = doLabelLine Or la.LabelType = doLabelPolyLine Then
        If bROIArea Then funROIResultString = Int(la.ROILength) & la.ROIDistanceUnits
    Else
        If la.LabelType = doLabelEllipse Or la.LabelType = doLabelPolygon Or _
          la.LabelType = doLabelRectangle Then
                If bROIArea Then funROIResultString = strROIArea & Int(la.ROIArea) & la.ROIDistanceUnits
                If bROIMean Then
                    If funROIResultString <> "" Then
                        funROIResultString = funROIResultString & vbCrLf & strROIMean
                        If blnSelectedImageIfColor = True Then
                            funROIResultString = funROIResultString & "0"
                        Else
                            funROIResultString = funROIResultString & Int(la.ROIMean)
                        End If
                    Else
                        funROIResultString = strROIMean & Int(la.ROIMean)
                    End If
                End If
                If bROIStandardDeviation Then       '显示标准差
                    If funROIResultString <> "" Then
                        funROIResultString = funROIResultString & vbCrLf & strROIStdDev
                        If blnSelectedImageIfColor = True Then
                            funROIResultString = funROIResultString & "0"
                        Else
                            funROIResultString = funROIResultString & Int(la.ROIStandardDeviation)
                        End If
                    Else
                        funROIResultString = strROIStdDev & Int(la.ROIStandardDeviation)
                    End If
                End If
                If bROILength Then          '显示周长
                    If funROIResultString <> "" Then
                        funROIResultString = funROIResultString & vbCrLf & strROILength & Int(la.ROILength)
                    Else
                        funROIResultString = strROILength & Int(la.ROILength)
                    End If
                End If
                If bROIMax Then             '显示最大值
                    If funROIResultString <> "" Then
                        funROIResultString = funROIResultString & vbCrLf & strROIMax & Int(la.ROIMax)
                    Else
                        funROIResultString = strROIMax & Int(la.ROIMax)
                    End If
                End If
                If bROIMin Then             '显示最小值
                    If funROIResultString <> "" Then
                        funROIResultString = funROIResultString & vbCrLf & strROIMin & Int(la.ROIMin)
                    Else
                        funROIResultString = strROIMin & Int(la.ROIMin)
                    End If
                End If
        End If
    End If
End Function

Public Sub subMove25(im As DicomImage, f As frmViewer)
'------------------------------------------------
'功能：裁减根据系统标注1，移动标注2-5。其中2－左边；3-下边；4-右边；5-上边。
'参数：im--需要移动系统标注的图像；f--需要移动系统标注的窗体
'返回：无，直接移动标注2,3,4,5的位置
'------------------------------------------------
    Dim i As Integer
    For i = 2 To 5
        '设置四个遮挡矩形的宽度和高度
        im.Labels(i).height = 32766 \ IIf(i Mod 2 = 0, 1, 2)
        im.Labels(i).width = 32766 \ IIf(i Mod 2 = 0, 2, 1)
    Next
    If im.Labels(1).width > 0 Then
        im.Labels(2).left = im.Labels(1).left - im.Labels(2).width
        im.Labels(4).left = im.Labels(1).left + im.Labels(1).width
    Else
        im.Labels(2).left = im.Labels(1).left + im.Labels(1).width - im.Labels(4).width
        im.Labels(4).left = im.Labels(1).left
    End If
    
    If im.Labels(1).height > 0 Then
        im.Labels(3).top = im.Labels(1).top + im.Labels(1).height
        im.Labels(5).top = im.Labels(1).top - im.Labels(5).height
    Else
        im.Labels(3).top = im.Labels(1).top
        im.Labels(5).top = im.Labels(1).top + im.Labels(1).height - im.Labels(5).height
    End If
    im.Labels(2).top = im.Labels(5).top
    im.Labels(4).top = im.Labels(5).top
    im.Labels(3).left = im.Labels(2).left
    im.Labels(5).left = im.Labels(2).left
End Sub

Public Sub subTakeOut1(ls As DicomLabels, im As DicomImage, isTakeOut1 As Boolean)
'------------------------------------------------
'功能：从选中的标注中去掉系统标注用于裁减的序号为2-5的标注，对于序号1的标注是否删除，通过参数isTakeOut1来判断。
'参数：ls--需要进行删除的标注集；im--标注所在的图像；isTakeOut1--是否删除序号为1的标注，True删除，Fasle不删除。
'返回：无，直接处理标注集ls的内容。
'------------------------------------------------
    Dim l As DicomLabel
    For Each l In ls
        If im.Labels.IndexOf(l) < 6 Then
            If isTakeOut1 Or im.Labels.IndexOf(l) <> 1 Then ls.Remove (ls.IndexOf(l))
        End If
    Next
End Sub

Public Function funLabelType(la As DicomLabel) As String
'------------------------------------------------
'功能：返回内部使用的标注类型名称
'参数：la--需要判断标注类型的标注。
'返回：标注类型的中文名称
'上级函数或过程：frmLabelObject.load
'下级函数或过程：无
'引用的外部参数：无
'编制人：胡涛
'------------------------------------------------
    funLabelType = ""
    Select Case la.LabelType
        Case 0:
            funLabelType = "文字"
        Case 1:
            funLabelType = "椭圆"
        Case 2:
            funLabelType = "矩形"
        Case 3:
            If la.Tag = "JD1" Then
                funLabelType = "角度线(1)"
            ElseIf la.Tag = "JD2" Then
                funLabelType = "角度线(2)"
            ElseIf la.Tag = "JDT" Then
                funLabelType = "角度文字"
            ElseIf la.Tag = "RLL" Then
                funLabelType = "定位线"
            Else
                funLabelType = "直线"
            End If
        Case 4:
            funLabelType = "多边形"
        Case 5:
            funLabelType = "多边线"
        Case 6:
            funLabelType = "图像"
        Case 7:
            funLabelType = "体位标注"
        Case 8:
             funLabelType = "圆弧"
        Case 9:
             funLabelType = "内插值算法的多边形"
        Case 10:
             funLabelType = "箭头"
        Case 11:
             funLabelType = "标尺"
    End Select
End Function

Public Sub SubDispPeriod(la As DicomLabel, im As DicomImage, f As frmViewer)
'------------------------------------------------
'功能：为指定图像中的指定标注，显示标注选择句柄
'参数：la--被标注选择句柄包围的标注；im--显示标注选择句柄的图像；f--显示标注选择句柄的窗体。
'返回：无，直接显示标注选择句柄。
'2009用
'------------------------------------------------
    Dim intZoom As Double
    Dim i As Integer        '做循环用的临时变量
    Dim img As DicomImage
    Dim lblTemp As DicomLabel
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''计算句柄的显示大小
    If im.ActualZoom <> 0 Then
        For i = 11 To 20
            im.Labels(i).height = IIf(intPeriodSize / im.ActualZoom >= 1, intPeriodSize / im.ActualZoom, 1)
            im.Labels(i).width = im.Labels(i).height
        Next
    End If
    SubNoDispPeriod im, f               ''''为指定图像隐藏标注选择句柄
    Set im.Labels(11).TagObject = la    ''''''就是用1号句柄指向当前标注''这是一个重要点,以后很多地方要用到此点的记录
    
    If la.LabelType = doLabelLine Or la.LabelType = doLabelArrow Then
     ''线和箭头类型标注的处理，包括直线、箭头、三角形、血管狭窄、心胸比测量。
     ''使用的选择句柄序号为:直线和箭头（11，15）,角度（11,15,18）
     ''血管狭窄测量（11,12,13,14,15,16,17，18)，心胸比（11,14,15,18）
        
        If la.Tag = "VAS1L" Or la.Tag = "VAS2L" Then    '对血管狭窄测量标注进行处理
            Set lblTemp = la
            Set im.Labels(11).TagObject = lblTemp    '用1号句柄指向垂直线
            '给垂直线加选择句柄
            SubDispLinePeriod lblTemp, im, 11, 14
            '给垂直线加选择句柄
            For i = 1 To 4
                If Not lblTemp.TagObject Is Nothing Then
                    Set lblTemp = lblTemp.TagObject
                End If
            Next i
            If Right(lblTemp.Tag, 1) = "L" Then
                SubDispLinePeriod lblTemp, im, 15, 18
            Else    '隐藏序号为15,18的选择句柄
                
            End If
        ElseIf la.Tag = "CTR1L" Or la.Tag = "CTR2L" Then  '对心胸比测量的标注进行处理
            Set lblTemp = la
            Set im.Labels(11).TagObject = lblTemp   '用1号句柄指向直线
            Call SubDispLinePeriod(lblTemp, im, 11, 14)
            If Not lblTemp.TagObject Is Nothing Then
                Set lblTemp = lblTemp.TagObject
                If Not lblTemp.TagObject Is Nothing Then
                    Set lblTemp = lblTemp.TagObject
                End If
            End If
            If Right(lblTemp.Tag, 1) = "L" Then
                Call SubDispLinePeriod(lblTemp, im, 15, 18)
            End If
        Else
            SubDispLinePeriod la, im, 11, 15
        End If
        
        '''''''''角度的处理'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Mid(la.Tag, 1, 3) = "JD1" Or Mid(la.Tag, 1, 3) = "JD2" Then
            Dim laTagObject As New DicomLabel
            If Mid(la.Tag, 1, 3) = "JD1" Then
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Set laTagObject = la.TagObject.TagObject
                im.Labels(18).left = (laTagObject.left + laTagObject.width)
                im.Labels(18).top = (laTagObject.top + laTagObject.height)
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If laTagObject.width > 0 And laTagObject.height > 0 Then
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.width > 0 And laTagObject.height < 0 Then
                    im.Labels(18).top = im.Labels(18).top - im.Labels(18).height
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.width < 0 And laTagObject.height < 0 Then
                    im.Labels(18).left = im.Labels(18).left - im.Labels(18).width
                    im.Labels(18).top = im.Labels(18).top - im.Labels(18).height
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.width < 0 And laTagObject.height > 0 Then
                    im.Labels(18).left = im.Labels(18).left - im.Labels(18).width
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.width = 0 Then  ''处理竖线
                    If laTagObject.height > 0 Then
                    Else
                        im.Labels(18).top = im.Labels(18).top - im.Labels(18).height
                    End If
                    im.Labels(18).left = im.Labels(18).left - im.Labels(18).width / 2
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.height = 0 Then   ''处理横线
                    If laTagObject.width > 0 Then
                    Else
                        im.Labels(18).left = im.Labels(18).left - im.Labels(18).height
                    End If
                    im.Labels(18).top = im.Labels(18).top - im.Labels(18).width / 2
                End If
            Else
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Set laTagObject = la.TagObject
                im.Labels(18).left = laTagObject.left
                im.Labels(18).top = laTagObject.top
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If laTagObject.width > 0 And laTagObject.height > 0 Then
                    im.Labels(18).left = im.Labels(18).left - im.Labels(18).width
                    im.Labels(18).top = im.Labels(18).top - im.Labels(18).height
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.width > 0 And laTagObject.height < 0 Then
                    im.Labels(18).left = laTagObject.left - im.Labels(18).width
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.width < 0 And laTagObject.height < 0 Then
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.width < 0 And laTagObject.height > 0 Then
                    im.Labels(18).top = im.Labels(18).top - im.Labels(18).height
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.width = 0 Then  ''处理竖线
                    If laTagObject.height > 0 Then
                        im.Labels(18).top = im.Labels(18).top - im.Labels(18).height
                    Else
                    End If
                    im.Labels(18).left = im.Labels(18).left - im.Labels(18).width / 2
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.height = 0 Then   ''处理横线
                    If laTagObject.width > 0 Then
                        im.Labels(18).left = im.Labels(18).left - im.Labels(18).height
                    Else
                    End If
                    im.Labels(18).top = im.Labels(18).top - im.Labels(18).width / 2
                End If
            End If
        End If          '角度处理结束
        
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 11 To 18
          im.Labels(i).Visible = True
        Next
    ElseIf la.LabelType = doLabelEllipse Or la.LabelType = doLabelRectangle Then
    '''''''''''''''''''''''''''''''''''' ''''矩形和椭圆''''''''''''''''''''''''''''''''''''''''''''''''''''''
        im.Labels(11).left = la.left
        im.Labels(11).top = la.top
        im.Labels(12).top = (la.top + (la.height - im.Labels(11).height) / 2)
        im.Labels(13).top = (la.top + la.height)
        im.Labels(14).left = (la.left + (la.width - im.Labels(11).height) / 2)
        im.Labels(15).left = (la.left + la.width)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If la.width > 0 And la.height > 0 Then
            im.Labels(11).left = im.Labels(11).left - im.Labels(11).width
            im.Labels(11).top = im.Labels(11).top - im.Labels(11).height
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ElseIf la.width > 0 And la.height < 0 Then
            im.Labels(11).left = la.left - im.Labels(11).width
            im.Labels(13).top = im.Labels(13).top - im.Labels(11).width
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ElseIf la.width < 0 And la.height < 0 Then
            im.Labels(13).left = im.Labels(13).left - im.Labels(11).width
            im.Labels(13).top = im.Labels(13).top - im.Labels(11).width
            im.Labels(15).left = im.Labels(15).left - im.Labels(11).width
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ElseIf la.width < 0 And la.height > 0 Then
            im.Labels(15).left = im.Labels(15).left - im.Labels(11).width
            im.Labels(13).left = im.Labels(13).left - im.Labels(11).width
            im.Labels(11).top = im.Labels(11).top - im.Labels(11).height
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        im.Labels(12).left = im.Labels(11).left
        im.Labels(13).left = im.Labels(11).left
        im.Labels(14).top = im.Labels(13).top
        im.Labels(15).top = im.Labels(13).top
        im.Labels(16).left = im.Labels(15).left
        im.Labels(16).top = im.Labels(12).top
        im.Labels(17).left = im.Labels(15).left
        im.Labels(17).top = im.Labels(11).top
        im.Labels(18).left = im.Labels(14).left
        im.Labels(18).top = im.Labels(11).top
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 11 To 18
          im.Labels(i).Visible = True
        Next
    ElseIf la.LabelType = 4 Or la.LabelType = 5 Then
        '''''''''''''''''''''''''''''''''''' ''''多边形和多边线'''''''''''''''''''''''''''''''''''''''''''''''''''''
        la.SelectMode = 4
    End If
End Sub

Private Sub SubDispLinePeriod(la As DicomLabel, im As DicomImage, intEnd1 As Integer, intEnd2 As Integer)
'------------------------------------------------
'功能：为图像中的指定直线和箭头标注，显示标注选择句柄
'参数：la--被标注选择句柄包围的标注；im--显示标注选择句柄的图像；intEnd1-第一个句柄序号；intEnd2-第二个句柄序号
'返回：无，直接显示标注选择句柄。
'2009用
'------------------------------------------------
    If la.LabelType = doLabelLine Or la.LabelType = doLabelArrow Then
        im.Labels(intEnd1).left = la.left
        im.Labels(intEnd1).top = la.top
        im.Labels(intEnd2).left = (la.left + la.width)
        im.Labels(intEnd2).top = (la.top + la.height)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If la.width > 0 And la.height > 0 Then
            im.Labels(intEnd1).left = im.Labels(intEnd1).left - im.Labels(intEnd1).width
            im.Labels(intEnd1).top = im.Labels(intEnd1).top - im.Labels(intEnd1).height
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ElseIf la.width > 0 And la.height < 0 Then
            im.Labels(intEnd1).left = la.left - im.Labels(intEnd1).width
            im.Labels(intEnd2).top = im.Labels(intEnd2).top - im.Labels(intEnd2).height
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ElseIf la.width < 0 And la.height < 0 Then
            im.Labels(intEnd2).left = im.Labels(intEnd2).left - im.Labels(intEnd2).width
            im.Labels(intEnd2).top = im.Labels(intEnd2).top - im.Labels(intEnd2).height
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ElseIf la.width < 0 And la.height > 0 Then
            im.Labels(intEnd1).top = im.Labels(intEnd1).top - im.Labels(intEnd1).height
            im.Labels(intEnd2).left = im.Labels(intEnd2).left - im.Labels(intEnd2).width
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ElseIf la.width = 0 Then  ''处理竖线
            If la.height > 0 Then
                im.Labels(intEnd1).top = im.Labels(intEnd1).top - im.Labels(intEnd1).height
            Else
                im.Labels(intEnd2).top = im.Labels(intEnd2).top - im.Labels(intEnd2).height
            End If
            im.Labels(intEnd1).left = im.Labels(intEnd1).left - im.Labels(intEnd1).width / 2
            im.Labels(intEnd2).left = im.Labels(intEnd2).left - im.Labels(intEnd2).width / 2
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ElseIf la.height = 0 Then   ''处理横线
            If la.width > 0 Then
                im.Labels(intEnd1).left = im.Labels(intEnd1).left - im.Labels(intEnd1).height
            Else
                im.Labels(intEnd2).left = im.Labels(intEnd2).left - im.Labels(intEnd2).height
            End If
            im.Labels(intEnd1).top = im.Labels(intEnd1).top - im.Labels(intEnd1).width / 2
            im.Labels(intEnd2).top = im.Labels(intEnd2).top - im.Labels(intEnd2).width / 2
        End If
    End If
End Sub

Public Sub subCutOutInphase(v As DicomViewer, im As DicomImage, f As frmViewer)
'------------------------------------------------
'功能：在裁减状态下对裁减操作作出图像同步处理
'参数：v--进行图像同步的viewer；im--做为同步参照的图像；f--进行同步的窗体
'返回：无，直接改变裁剪标注的位置和大小
'------------------------------------------------
    Dim img As DicomImage
    Dim i As Integer
    For Each img In v.Images
        For i = 1 To 5
            img.Labels(i).Visible = im.Labels(i).Visible
            img.Labels(i).left = im.Labels(i).left
            img.Labels(i).top = im.Labels(i).top
            img.Labels(i).width = im.Labels(i).width
            img.Labels(i).height = im.Labels(i).height
        Next
        SubNoDispPeriod img, f          '为指定图像隐藏标注选择句柄
        If img.Labels(1).Visible Then SubDispPeriod img.Labels(1), img, f   '为指定图像中的指定标注，显示标注选择句柄
    Next
    v.Refresh
End Sub

Public Function funMouseOverPeriod(v As DicomViewer, im As DicomImage, ByVal x As Long, ByVal y As Long) As Integer
'------------------------------------------------
'功能：返回鼠标所越过的句柄编号
'参数：v--鼠标所在的viewer；im--鼠标所在的图像；x--鼠标的X位置；y--鼠标的Y位置。
'返回：0--鼠标不在句柄上；11到18-鼠标在该序号所代表的句柄上。
'2009用
'------------------------------------------------
    Dim xx As Long, Yy As Long
    xx = v.ImageXPosition(x, y)
    Yy = v.ImageYPosition(x, y)
    funMouseOverPeriod = 0
    Dim i As Integer
    With im
        For i = 11 To 18
            If .Labels(i).Visible And .Labels(i).left <= xx And .Labels(i).top <= Yy And .Labels(i).top + .Labels(i).height >= Yy And .Labels(i).left + .Labels(i).width >= xx Then
                funMouseOverPeriod = i
                Exit For
            End If
        Next
    End With
End Function

Private Sub subMoveMPRLabel(f As frmViewer, la As DicomLabel, xx As Integer, Yy As Integer, basex As Long, baseY As Long)
'------------------------------------------------
'功能：移动矢冠状重建控制点、线，且生成新的重建图像。
'参数： f--进行矢冠状重建的窗体；
'       la--被移动的矢冠状重建控制点或控制线；
'       xx --标注新位置在图像上的X坐标；
'       yy --标注新位置在图像上的Y坐标；
'       basex--旧位置的图像像素x坐标；
'       baseY--旧位置的图像像素y坐标。
'返回：无，直接移动矢冠状重建的控制点和线，并生成重建结果图像。
'2009用
'------------------------------------------------
    Dim intIndex As Integer
    
    On Error GoTo err
    
    ''''''''''''''''''''''[是四角点的移动]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    intIndex = f.SelectedImage.Labels.IndexOf(la)
    '矢冠状控制点中四个边点的移动处理
    If intIndex = G_INT_SYS_LABEL_MPR_POINT_V1 Then
        Call subPeriodMove(f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V1), xx, Yy, _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V2), _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPRV), _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), f.SelectedImage)
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_POINT_H1 Then
        Call subPeriodMove(f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H1), xx, Yy, _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H2), _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPRH), _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), f.SelectedImage)
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_POINT_V2 Then
        Call subPeriodMove(f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V2), xx, Yy, _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V1), _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPRV), _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), f.SelectedImage)
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_POINT_H2 Then
        Call subPeriodMove(f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H2), xx, Yy, _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H1), _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPRH), _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), f.SelectedImage)
    '''''''''''''''''''''''''''中心点的移动'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_POINT_O Then
        f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top = f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top + Yy - baseY
        f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left = f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left + xx - basex
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top < -G_INT_MPR_RADIUS / 2 + 1 Then
            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top = -G_INT_MPR_RADIUS / 2 + 1
        End If
        If f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left < -G_INT_MPR_RADIUS / 2 + 1 Then
            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left = -G_INT_MPR_RADIUS / 2 + 1
        End If
        If f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top > f.SelectedImage.sizey - G_INT_MPR_RADIUS / 2 - 1 Then
            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top = f.SelectedImage.sizey - G_INT_MPR_RADIUS - 1
        End If
        If f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left > f.SelectedImage.sizex - G_INT_MPR_RADIUS / 2 - 1 Then
            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left = f.SelectedImage.sizex - G_INT_MPR_RADIUS - 1
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '矢冠状中心点的移动
        If xx <> basex Then
            Call subPeriodMovee5X(f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V1), _
                                f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V2), _
                                f.SelectedImage.Labels(G_INT_SYS_LABEL_MPRV), xx, Yy, _
                                f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), basex, baseY, f.SelectedImage)
        End If
        
        If Yy <> baseY Then
            Call subPeriodMovee5X(f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H1), _
                                f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H2), _
                                f.SelectedImage.Labels(G_INT_SYS_LABEL_MPRH), xx, Yy, _
                                f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), basex, baseY, f.SelectedImage)
        End If
        
    End If
    
    ''''''''''进行重建''''''''''''''''''''''''''''''''''''''''''''''
    '标注是MPR控制线竖线的两个端点，或者是MPR控制线竖线的中心点，此时需要移动的是MPR竖线
    If intIndex = G_INT_SYS_LABEL_MPR_POINT_V1 Or intIndex = G_INT_SYS_LABEL_MPR_POINT_V2 _
        Or (intIndex = G_INT_SYS_LABEL_MPR_POINT_O And xx <> basex) Then
        If funGetMPRImageAndShow(f.SelectedImage.Labels(G_INT_SYS_LABEL_MPRV), f, _
                                    f.Viewer(ZLMPRCube(1).intViewerIndex), f.SelectedImageIndex, _
                                    ZLMPRCube(2).intViewerIndex, ToltalHeight, 1, False, True) = False Then
            Call funMPR(f, True)
            Exit Sub
        End If
    End If
    
    '标注是MPR控制线横线的两个端点，或者是MPR控制线的横线中心点，此时需要移动的是MPR横线
    If intIndex = G_INT_SYS_LABEL_MPR_POINT_H1 Or intIndex = G_INT_SYS_LABEL_MPR_POINT_H2 _
        Or (intIndex = G_INT_SYS_LABEL_MPR_POINT_O And Yy <> baseY) Then
        If funGetMPRImageAndShow(f.SelectedImage.Labels(G_INT_SYS_LABEL_MPRH), f, _
                                    f.Viewer(ZLMPRCube(1).intViewerIndex), f.SelectedImageIndex, _
                                    ZLMPRCube(3).intViewerIndex, ToltalHeight, 2, False, True) = False Then
            Call funMPR(f, True)
            Exit Sub
        End If
    End If
    '''''''''矢冠状点线同步'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    subMPRLinenPhase f.Viewer(f.intSelectedSerial), f.SelectedImage
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub subMoveMPRReslutLabel(f As frmViewer, la As DicomLabel, x As Long, y As Long)
'------------------------------------------------
'功能：移动矢冠状重建结果线，横线控制轴位图像的自动翻页，竖线控制结果图
'参数： f--进行矢冠状重建的窗体；
'       la--被移动的矢冠状重建控制点或控制线；
'       x--x方向移动的图像像素距离；
'       y--y方向移动的图像像素距离；
'返回：无，直接移动矢冠状重建的结果线，在轴位图像翻页。
'------------------------------------------------
    Dim iImageIndex As Integer
    Dim OldIntSelectedSeries As Integer
    Dim OldSelectedImage As DicomImage
    Dim oldSelectedImageIndex As Integer
    Dim intIndex As Integer
    Dim lngNewPosLeft As Long   '新位置LEFT
    Dim lngNewPosTop As Long    '新位置TOP
    Dim dblH As Double
    Dim dblW As Double
    Dim iViewerIndex As Integer
    Dim OldLeft As Long
    Dim dblXieBian As Double
    Dim dblSin As Double
    Dim dblCos As Double
    Dim dblDistance As Double
    
    On Error GoTo err
    
    intIndex = f.SelectedImage.Labels.IndexOf(la)

    
    If intIndex = G_INT_SYS_LABEL_MPR_RESULT_H Then '横线，是翻页
        '先移动矢冠状重建结果线
        la.top = la.top + y
        
        '确保结果线不会离开图像
        If la.top < 0 Then
            la.top = 0
        ElseIf la.top > f.SelectedImage.sizey Then
            la.top = f.SelectedImage.sizey
        End If
        
        '根据结果线的位置，翻页
        iImageIndex = la.top / f.SelectedImage.sizey * f.Viewer(ZLMPRCube(1).intViewerIndex).Images.Count
        If iImageIndex > 0 And iImageIndex <= f.Viewer(ZLMPRCube(1).intViewerIndex).Images.Count Then
            OldIntSelectedSeries = f.intSelectedSerial
            Set OldSelectedImage = f.SelectedImage
            
            f.VScro(ZLMPRCube(1).intViewerIndex).Value = iImageIndex
            
            '原图翻页后，基本参数会改变，需要恢复
            f.intSelectedSerial = OldIntSelectedSeries
            Set f.SelectedImage = OldSelectedImage
        End If
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_RESULT_V Then '竖线，是重建图像
        OldLeft = la.left
        '先移动矢冠状重建结果线
        la.left = la.left + x
        
        '确保结果线不会离开图像
        If la.left < 0 Then
            la.left = 0
        ElseIf la.left > f.SelectedImage.sizex Then
            la.left = f.SelectedImage.sizex
        End If
        
        '根据结果线的位置，移动轴位图的对应控制线
        
        '计算新位置
        iViewerIndex = ZLMPRCube(1).intViewerIndex
        iImageIndex = f.VScro(iViewerIndex).Value
        
        '先除以做平方之后，再乘以10，否则平方时会溢出
        '先找到当前的图像是第二幅图，还是第三幅图
        If f.intSelectedSerial = ZLMPRCube(2).intViewerIndex Then
            '第二幅图，移动轴位图中的竖线
            dblH = f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPRV).height / 10
            dblW = f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPRV).width / 10
        ElseIf f.intSelectedSerial = ZLMPRCube(3).intViewerIndex Then
            '第三幅图，移动轴位图中的竖线
            dblH = f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPRH).height / 10
            dblW = f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPRH).width / 10
        End If
        
        dblXieBian = Sqr((dblW * dblW) + (dblH * dblH)) * 10
        dblSin = dblH * 10 / dblXieBian
        dblCos = dblW * 10 / dblXieBian
        dblDistance = (la.left - OldLeft) / f.SelectedImage.sizex * dblXieBian
        lngNewPosLeft = f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPR_POINT_O).left + dblDistance * dblCos
        lngNewPosTop = f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPR_POINT_O).top + dblDistance * dblSin
            
        '记录图像参数，将当前图像设置成轴位图像
        OldIntSelectedSeries = f.intSelectedSerial
        oldSelectedImageIndex = f.SelectedImageIndex
        Set OldSelectedImage = f.SelectedImage
        
        f.intSelectedSerial = iViewerIndex
        Set f.SelectedImage = f.Viewer(iViewerIndex).Images(iImageIndex)
        f.SelectedImageIndex = iImageIndex
        
        Call subMoveMPRLabel(f, f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPR_POINT_O), _
            CInt(lngNewPosLeft), CInt(lngNewPosTop), _
            f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPR_POINT_O).left, _
            f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPR_POINT_O).top)
            
        '原图后，基本参数会改变，需要恢复
        f.intSelectedSerial = OldIntSelectedSeries
        f.SelectedImageIndex = oldSelectedImageIndex
        Set f.SelectedImage = OldSelectedImage
        
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subMPRChanegImage(thisForm As frmViewer)
'MPR矢冠状位重建时，改变了当前选中的图像
'参数： thisForm --- MPR所在的窗体
    
    On Error GoTo err
    
    If thisForm.SelectedImage Is Nothing Then Exit Sub
    If thisForm.intSelectedSerial <> ZLMPRCube(1).intViewerIndex Then Exit Sub
    
    '重新显示MPR控制线-竖线对应的结果图
    If funGetMPRImageAndShow(thisForm.SelectedImage.Labels(G_INT_SYS_LABEL_MPRV), thisForm, _
                                thisForm.Viewer(ZLMPRCube(1).intViewerIndex), thisForm.SelectedImageIndex, _
                                ZLMPRCube(2).intViewerIndex, ToltalHeight, 1, False, False) = False Then
        Call funMPR(thisForm, True)
        Exit Sub
    End If
    
    '重新显示MPR控制线-横线对应的结果图
    If funGetMPRImageAndShow(thisForm.SelectedImage.Labels(G_INT_SYS_LABEL_MPRH), thisForm, _
                                thisForm.Viewer(ZLMPRCube(1).intViewerIndex), thisForm.SelectedImageIndex, _
                                ZLMPRCube(3).intViewerIndex, ToltalHeight, 2, False, False) = False Then
        Call funMPR(thisForm, True)
        Exit Sub
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subMoveLable(la As DicomLabel, x As Long, y As Long, f As frmViewer, xxx As Long, yyy As Long, basex As Long, baseY As Long)
'------------------------------------------------
'功能：移动一个标注,包括矢冠状重建标注、心胸比测量标注、用户标注和裁剪标注
'参数：la--被移动的标注；x--x方向移动的图像像素距离；y--y方向移动的图像像素距离；f--移动标注的窗体；
'      xxx--新位置的屏幕像素x坐标；yyy--新位置的屏幕像素y坐标；basex--旧位置的图像像素x坐标；
'      baseY--旧位置的图像像素y坐标。
'返回：无
'2009用
'------------------------------------------------
    Dim aa As Variant
    Dim lat As DicomLabel
    Dim i As Integer
    Dim pyX, pyY As Integer
    Dim lblTemp As DicomLabel
    
    
    If f.SelectedImage.Labels.IndexOf(la) >= G_INT_SYS_LABEL_MPRV And f.SelectedImage.Labels.IndexOf(la) <= G_INT_SYS_LABEL_MPR_POINT_O Then ''[矢冠状线的移动]
        '移动矢冠状重建控制点、线，且生成新的重建图像。
        Dim xx As Integer
        Dim Yy As Integer
        
        xx = f.Viewer(f.intSelectedSerial).ImageXPosition(xxx, yyy)
        Yy = f.Viewer(f.intSelectedSerial).ImageYPosition(xxx, yyy)
        
        Call subMoveMPRLabel(f, la, xx, Yy, basex, baseY)
    ElseIf (f.SelectedImage.Labels.IndexOf(la) = G_INT_SYS_LABEL_MPR_RESULT_H) _
        Or (f.SelectedImage.Labels.IndexOf(la) = G_INT_SYS_LABEL_MPR_RESULT_V) Then
        '移动矢冠状重建结果线
        Call subMoveMPRReslutLabel(f, la, x, y)
    Else                                '用户标注的移动
        la.left = la.left + x
        la.top = la.top + y
        
        If la.LabelType = 4 Or la.LabelType = 5 Then    '处理多边形和多边线
            aa = la.Points
             For i = 1 To UBound(aa) Step 2
                 aa(i) = aa(i) + x
                 aa(i + 1) = aa(i + 1) + y
             Next
            la.Points = aa
            If la.LabelType = doLabelPolygon And Not la.TagObject Is Nothing Then la.TagObject.Text = funROIResultString(la, f.SelectedImage)
        End If
        ''''''''''''''''''对于角度线的处理'''
        If Mid(la.Tag, 1, 2) = "JD" And Mid(la.Tag, 1, 3) <> "JDT" Then
            la.TagObject.left = la.TagObject.left + x
            la.TagObject.top = la.TagObject.top + y
            la.TagObject.TagObject.left = la.TagObject.TagObject.left + x
            la.TagObject.TagObject.top = la.TagObject.TagObject.top + y
            If Mid(la.Tag, 1, 3) = "JD1" Then
                la.TagObject.AnchorX = la.left '
                la.TagObject.AnchorY = la.top '
            Else
                la.TagObject.TagObject.AnchorX = la.TagObject.left '
                la.TagObject.TagObject.AnchorY = la.TagObject.top '
            End If
        ElseIf left(la.Tag, 3) = "VAS" Then         '处理“血管狭窄测量"
            Dim iVasCount As Integer        '记录血管狭窄测量标注的数量，全部标注有8个，只画正常标注有4个
            '先恢复la的位置
            iVasCount = 8
            If la.Tag = la.TagObject.TagObject.TagObject.TagObject.Tag Then iVasCount = 4
            la.left = la.left - x
            la.top = la.top - y
            '移动剩下的7个标注
            Set lblTemp = la
            If lblTemp.Tag = "VAS1L" Or lblTemp.Tag = "VAS2L" Then
                Set lblTemp = lblTemp.TagObject.TagObject.TagObject
            ElseIf lblTemp.Tag = "VAS1T" Or lblTemp.Tag = "VAS2T" Then
                Set lblTemp = lblTemp.TagObject.TagObject
            ElseIf lblTemp.Tag = "VAS1E1" Or lblTemp.Tag = "VAS2E1" Then
                Set lblTemp = lblTemp.TagObject
            End If
            For i = 1 To iVasCount
                Set lblTemp = lblTemp.TagObject
                lblTemp.left = lblTemp.left + x
                lblTemp.top = lblTemp.top + y
                If lblTemp.Tag = "VAS1L" Or lblTemp.Tag = "VAS2L" Then
                    lblTemp.TagObject.AnchorX = lblTemp.left + lblTemp.width / 2
                    lblTemp.TagObject.AnchorY = lblTemp.top + lblTemp.height / 2
                End If
                If Mid(lblTemp.Tag, 5, 1) = "E" Then
                    lblTemp.Text = Val(left(lblTemp.Text, InStr(lblTemp.Text, ",") - 1)) + x & "," _
                                   & Val(Right(lblTemp.Text, Len(lblTemp.Text) - InStr(lblTemp.Text, ","))) + y
                End If
            Next i
        ElseIf left(la.Tag, 3) = "CTR" Then     '处理心胸比测量标注
            Dim iCtrCount As Integer    '记录心胸比测量标注的数量，画完全部标注有4个，只画了心脏线则只有2个
            iCtrCount = 4
            If la.Tag = la.TagObject.TagObject.Tag Then iCtrCount = 2
            
            la.TagObject.AnchorX = la.left + la.width / 2
            la.TagObject.AnchorY = la.top + la.height / 2
            Set lblTemp = la
            For i = 1 To iCtrCount - 1
                Set lblTemp = lblTemp.TagObject
                lblTemp.left = lblTemp.left + x
                lblTemp.top = lblTemp.top + y
                If Right(lblTemp.Tag, 1) = "L" Then     '文字标注指向L标注
                    lblTemp.TagObject.AnchorX = lblTemp.left + lblTemp.width / 2
                    lblTemp.TagObject.AnchorY = lblTemp.top + lblTemp.height / 2
                End If
            Next i
        Else            '其他标注的处理，文字、箭头、直线、区域、椭圆、矩形等
            If la.LabelType <> doLabelText Then         ''不是文字的处理
                If Not la.TagObject Is Nothing Then
                    Set lat = la.TagObject
                    If la.LabelType <> doLabelArrow Then        '''''不是箭头
                        lat.AnchorX = la.left + la.width / 2
                        lat.AnchorY = la.top + la.height / 2
                        If la.LabelType = doLabelLine Or la.LabelType = doLabelPolyLine Then
                            lat.Text = funROIResultString(la, f.SelectedImage)
                        Else
                            lat.Text = ""
                        End If
                    Else
                        lat.AnchorX = la.left + la.width
                        lat.AnchorY = la.top + la.height
                    End If
                    lat.left = lat.left + x
                    lat.top = lat.top + y
                End If
            End If
        End If
        ''''''''''如果是裁减矩形框则''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If f.SelectedImage.Labels.IndexOf(la) = 1 Then
            subMove25 f.SelectedImage, f        '裁减根据系统标注1 移动2-4
            If Button_miImageInPhase = True Then subCutOutInphase f.Viewer(f.intSelectedSerial), f.SelectedImage, f
        End If
    End If
End Sub

Public Sub subChangeLableSize(la As DicomLabel, x As Long, y As Long, iR As Integer, f As frmViewer)
'------------------------------------------------
'功能：改变一个标注的大小,并修改其相关测量信息的显示值。
'参数：la--需要改变大小的标注；x--标注X方向移动的距离；y--标注Y方向移动的距离；
'      iR--标记通过哪个句柄来移动标注，对11到18号句柄分别有不同的处理方法，11-18号句柄分别表示左上、左中、左下、下中、右下、右中、右上，上中。
'      f--移动标注的窗体。
'返回：无，直接修改标注。
'2009用
'------------------------------------------------
    Dim lat As DicomLabel
    Dim lblTemp As DicomLabel
    '移动标注的位置
    '血管狭窄测量使用标注(11，14)--VAS1L，(15，18)--VAS2L，关系是VAS1L-VAS1T-VAS1E1-VAS1E2-VAS2L-VAS2T-VAS2E1-VAS2E2
    '心胸比测量使用标注(11,14)--CTR1L,(15,18)--CTR2L,关系是CTR1L-CTR1T-CTR2L-CTR2T
    '角度测量使用标注（11,15,18）,关系是JD2-JD1-JDT
    
    If iR = 11 Then         '左上角的句柄
        la.left = la.left + x
        la.width = la.width - x
        la.top = la.top + y
        la.height = la.height - y
        If Mid(la.Tag, 1, 3) = "JD2" Then       '处理角度标注
            la.TagObject.width = la.TagObject.width + x
            la.TagObject.height = la.TagObject.height + y
        End If
    ElseIf iR = 12 Then     '左中的句柄
        la.left = la.left + x
        la.width = la.width - x
    ElseIf iR = 13 Then     '左下角的句柄
        la.left = la.left + x
        la.width = la.width - x
        la.height = la.height + y
    ElseIf iR = 14 Then     '下中的句柄
        If left(la.Tag, 3) = "VAS" Then     '血管狭窄测量
            la.height = la.height + y
            la.width = la.width + x
        ElseIf left(la.Tag, 3) = "CTR" Then '心胸比测量
            la.height = la.height + y
            la.width = la.width + x
        Else    '其他情况，矩形或圆形
            la.height = la.height + y
        End If
    ElseIf iR = 15 Then     '右下角的句柄
        If Mid(la.Tag, 1, 3) = "JD1" Then   '处理角度标注
            la.width = la.width + x
            la.height = la.height + y
            la.TagObject.TagObject.left = la.TagObject.TagObject.left + x
            la.TagObject.TagObject.width = la.TagObject.TagObject.width - x
            la.TagObject.TagObject.top = la.TagObject.TagObject.top + y
            la.TagObject.TagObject.height = la.TagObject.TagObject.height - y
        ElseIf left(la.Tag, 3) = "VAS" Then '血管狭窄测量
            Set lblTemp = la.TagObject.TagObject.TagObject.TagObject
            lblTemp.left = lblTemp.left + x
            lblTemp.width = lblTemp.width - x
            lblTemp.top = lblTemp.top + y
            lblTemp.height = lblTemp.height - y
        ElseIf left(la.Tag, 3) = "CTR" Then '心胸比测量，需要专门处理文字标注的位置
            Set lblTemp = la.TagObject.TagObject
            lblTemp.left = lblTemp.left + x
            lblTemp.width = lblTemp.width - x
            lblTemp.top = lblTemp.top + y
            lblTemp.height = lblTemp.height - y
        Else                                '处理其他标注
            la.width = la.width + x
            la.height = la.height + y
        End If
    ElseIf iR = 16 Then     '右中的句柄
        la.width = la.width + x
    ElseIf iR = 17 Then     '右上角的句柄
        la.top = la.top + y
        la.height = la.height - y
        la.width = la.width + x
    ElseIf iR = 18 Then     '上中的句柄
        If Mid(la.Tag, 1, 2) = "JD" Then        '处理角度标注
            If Mid(la.Tag, 1, 3) = "JD1" Then
                la.TagObject.TagObject.width = la.TagObject.TagObject.width + x
                la.TagObject.TagObject.height = la.TagObject.TagObject.height + y
            Else
                '处理JD1的线
                la.TagObject.left = la.TagObject.left + x
                la.TagObject.width = la.TagObject.width - x
                la.TagObject.top = la.TagObject.top + y
                la.TagObject.height = la.TagObject.height - y
            End If
        ElseIf left(la.Tag, 3) = "VAS" Then     '血管狭窄测量
            Set lblTemp = la.TagObject.TagObject.TagObject.TagObject
            lblTemp.height = lblTemp.height + y
            lblTemp.width = lblTemp.width + x
        ElseIf left(la.Tag, 3) = "CTR" Then     '心胸比测量
            Set lblTemp = la.TagObject.TagObject
            lblTemp.height = lblTemp.height + y
            lblTemp.width = lblTemp.width + x
        Else                                    '处理其他标注
            la.top = la.top + y
            la.height = la.height - y
        End If
    End If
    
    '处理跟当前标注相连的其他配套标注，如关联文字信息等的位置和显示值
    '如果被选中的标注是裁剪标注，则进行相应处理
    If f.SelectedImage.Labels.IndexOf(la) = 1 Then
        subMove25 f.SelectedImage, f            '裁减根据系统标注1，移动标注2-5
        '在裁减状态下对裁减操作作出图像同步处理
        If Button_miImageInPhase = True Then subCutOutInphase f.Viewer(f.intSelectedSerial), f.SelectedImage, f
    Else
        If Mid(la.Tag, 1, 2) = "JD" Then           '处理角度标注
            If Mid(la.Tag, 1, 3) = "JD1" Then
                Set lat = la.TagObject
                la.TagObject.left = la.left
                la.TagObject.top = la.top
                la.TagObject.AnchorX = la.left
                la.TagObject.AnchorY = la.top
                f.lblChange = funROIResultString(la, f.SelectedImage)
                lat.Text = f.lblChange
            Else
                Set lat = la.TagObject.TagObject
                la.TagObject.TagObject.left = la.TagObject.left
                la.TagObject.TagObject.top = la.TagObject.top
                la.TagObject.TagObject.AnchorX = la.TagObject.left
                la.TagObject.TagObject.AnchorY = la.TagObject.top
                f.lblChange = funROIResultString(la, f.SelectedImage)
                lat.Text = f.lblChange
            End If
        ElseIf left(la.Tag, 3) = "VAS" Then     '处理血管狭窄测量
            If iR = 15 Or iR = 18 Then
                Set lblTemp = la.TagObject.TagObject.TagObject.TagObject
            Else
                Set lblTemp = la
            End If
            
            lblTemp.TagObject.left = lblTemp.left + lblTemp.width + intTextoOffX
            lblTemp.TagObject.top = lblTemp.top + lblTemp.height + intTextoOffY
            lblTemp.TagObject.AnchorX = lblTemp.left + lblTemp.width / 2
            lblTemp.TagObject.AnchorY = lblTemp.top + lblTemp.height / 2
        
            Call funDrawVas(lblTemp, f.SelectedImage, IIf(lblTemp.Tag = "VAS1L", 1, 2))
        ElseIf left(la.Tag, 3) = "CTR" Then     '处理心胸比测量
            If iR = 15 Or iR = 18 Then
                Set lblTemp = la.TagObject.TagObject
            Else
                Set lblTemp = la
            End If
            lblTemp.TagObject.left = lblTemp.left + lblTemp.width + intTextoOffX
            lblTemp.TagObject.top = lblTemp.top + lblTemp.height + intTextoOffY
            lblTemp.TagObject.AnchorX = lblTemp.left + lblTemp.width / 2
            lblTemp.TagObject.AnchorY = lblTemp.top + lblTemp.height / 2
            
            Call funcGetCadioThoracicRatio(la, f.SelectedImage)
        Else                                    '处理其他标注
            Set lat = la.TagObject
            If la.LabelType = doLabelArrow Then   '''箭头
                lat.left = la.left + la.width
                lat.top = la.top + la.height
                lat.AnchorX = la.left + la.width
                lat.AnchorY = la.top + la.height
            Else                '直线、椭圆、矩形、区域、曲线等标注
                '对非封闭区域标注，生成测量结果字符串,根据系统配置中设置的是否显示面积，平均值，均方差等条件
                If la.LabelType = doLabelEllipse Or la.LabelType = doLabelPolygon Or la.LabelType = doLabelRectangle Then
                    lat.Text = ""
                Else
                    lat.Text = funROIResultString(la, f.SelectedImage)
                End If
                lat.left = la.left + la.width + intTextoOffX
                lat.top = la.top + la.height + intTextoOffY
                lat.AnchorX = la.left + la.width / 2
                lat.AnchorY = la.top + la.height / 2
            End If
        End If
    End If
End Sub

Sub SubNoDispPeriod(im As DicomImage, f As frmViewer)
'------------------------------------------------
'功能：为指定图像隐藏标注选择句柄
'参数：im--需要隐藏标注选择句柄的图像；f--隐藏标注选择句柄的窗体
'返回：无，直接隐藏标注选择句柄
'2009用
'------------------------------------------------
    Dim i As Integer
    For i = 11 To 20
      im.Labels(i).Visible = False
      im.Labels(i).left = G_INT_SYS_LABEL_HIDE_LEFT
    Next
    im.Refresh False
    If Not f.DLblOld Is Nothing Then f.DLblOld.SelectMode = doSelectNone
End Sub

Public Sub subTextCoordinate(im As DicomImage, x, y, lb As Label)
'------------------------------------------------
'功能：根据图像的反转情况决定文字的坐标转换
'参数：im--发生翻转的图像；x--   y--   lb--
'返回：
'2009用
'------------------------------------------------
    Dim xx As Long, Yy As Long
    Dim TXY As Single
    TXY = im.sizex / im.sizey
    xx = x
    Yy = y
    
    '对于旋转的情况，重新计算新的x,y坐标
    If im.RotateState = doRotateNormal Then         '正常
        '不用处理
    ElseIf im.RotateState = doRotateLeft Then       '左传90度
        x = Yy
        y = im.sizey * TXY - xx - lb.height / Screen.TwipsPerPixelX / im.ActualZoom
    ElseIf im.RotateState = doRotate180 Then        '旋转180度
        x = im.sizex - xx - lb.width / Screen.TwipsPerPixelX / im.ActualZoom
        y = im.sizey - Yy - lb.height / Screen.TwipsPerPixelY / im.ActualZoom
    ElseIf im.RotateState = doRotateRight Then      '右转90度
        x = im.sizex / TXY - Yy - lb.width / Screen.TwipsPerPixelX / im.ActualZoom
        y = xx
    End If
    
    '处理左右镜象和上下倒置的情况，重新计算x,y坐标
    If im.FlipState = 1 Then            '左右镜象
        If im.RotateState = 0 Or im.RotateState = doRotate180 Then
             x = im.sizex - x - lb.width / Screen.TwipsPerPixelX / im.ActualZoom
        Else
             y = im.sizey * TXY - y - lb.height / Screen.TwipsPerPixelY / im.ActualZoom
        End If
    ElseIf im.FlipState = 2 Then        '上下倒置
        If im.RotateState = 0 Or im.RotateState = doRotate180 Then
             y = im.sizey - y - lb.height / Screen.TwipsPerPixelY / im.ActualZoom
        Else
            x = im.sizex / TXY - x - lb.height / Screen.TwipsPerPixelY / im.ActualZoom
        End If
    ElseIf im.FlipState = 3 Then        '左右镜象加上下倒置
        If im.RotateState = 0 Or im.RotateState = doRotate180 Then
            x = im.sizex - x - lb.width / Screen.TwipsPerPixelX / im.ActualZoom
            y = im.sizey - y - lb.height / Screen.TwipsPerPixelY / im.ActualZoom
        Else
            x = im.sizex / TXY - x - lb.width / Screen.TwipsPerPixelX / im.ActualZoom
            y = im.sizey * TXY - y - lb.height / Screen.TwipsPerPixelY / im.ActualZoom
        End If
    End If
End Sub

Public Sub SubChangeColor(la As DicomLabel, f As frmViewer)
'------------------------------------------------
'功能：改变选中LABEL的颜色
'参数：la--需要改变颜色的标注；f--改变标注颜色的窗体。
'返回：无，直接修改标注的颜色。
'2009用
'------------------------------------------------
    Dim lblTemp As DicomLabel
    Dim i As Integer
    '''''''''''''''''''''''''''[先恢复上一个被选中标注的颜色]'''''''''''''''''''''''''''''
    If Not f.DLblOld Is Nothing Then
        f.DLblOld.ForeColour = f.LngOldColor
        If Mid(f.DLblOld.Tag, 1, 2) = "JD" Then    ''如果是角度线的处理
            f.DLblOld.TagObject.ForeColour = f.LngOldColor
            If Not f.DLblOld.TagObject.TagObject Is Nothing Then f.DLblOld.TagObject.TagObject.ForeColour = f.LngOldColor
        ElseIf left(f.DLblOld.Tag, 3) = "VAS" Then      ''处理血管狭窄测量标注
            Set lblTemp = f.DLblOld
            For i = 1 To 7
                If Not lblTemp.TagObject Is Nothing Then
                    Set lblTemp = lblTemp.TagObject
                    lblTemp.ForeColour = f.LngOldColor
                End If
            Next i
        ElseIf left(f.DLblOld.Tag, 3) = "CTR" Then      ''处理心胸比测量标注
            Set lblTemp = f.DLblOld
            For i = 1 To 3
                If Not lblTemp.TagObject Is Nothing Then
                    Set lblTemp = lblTemp.TagObject
                    lblTemp.ForeColour = f.LngOldColor
                End If
            Next i
        Else
            If Not f.DLblOld.TagObject Is Nothing Then f.DLblOld.TagObject.ForeColour = f.LngOldColor
        End If
    End If
    '''''''''''''''''''''''''''[记录当前标注]'''''''''''''''''''''''''''''
    f.LngOldColor = la.ForeColour
    Set f.DLblOld = la
    '''''''''''''''''''''''''''''[改变当前被选中标注的颜色]'''''''''''''''''''''''''''
    la.ForeColour = lngLabelSelectedColor
    If la.LabelType <> doLabelText Then
        If Mid(la.Tag, 1, 2) = "JD" Then
            la.TagObject.ForeColour = lngLabelSelectedColor
            If Not la.TagObject.TagObject Is Nothing Then la.TagObject.TagObject.ForeColour = lngLabelSelectedColor
        ElseIf left(la.Tag, 3) = "VAS" Then
            Set lblTemp = la
            For i = 1 To 7
                If Not lblTemp.TagObject Is Nothing Then
                    Set lblTemp = lblTemp.TagObject
                    lblTemp.ForeColour = lngLabelSelectedColor
                End If
            Next i
        ElseIf left(la.Tag, 3) = "CTR" Then
            Set lblTemp = la
            For i = 1 To 3
                If Not lblTemp.TagObject Is Nothing Then
                    Set lblTemp = lblTemp.TagObject
                    lblTemp.ForeColour = lngLabelSelectedColor
                End If
            Next i
        Else
            If Not f.DLblOld.TagObject Is Nothing Then la.TagObject.ForeColour = lngLabelSelectedColor
        End If
    End If
End Sub

Public Function GetNewLabel(lType As Integer, lLeft As Integer, lTop As Integer, lWidth As Integer, lHeight As Integer) As DicomLabel
'------------------------------------------------
'功能：生成一个LABEL对象，并对其做初始化。
'参数：lType--标注的类型；lLeft--标注的Left值；lTop--标注的Top值；lWidth--标注的Width值；lHeight--标注的Height值。
'返回：新生成的标注。
'2009用
'------------------------------------------------
    Dim l As New DicomLabel
    l.LabelType = lType
    l.Transparent = True
    l.XOR = True
    l.ImageTied = True
    l.left = lLeft
    l.top = lTop
    l.width = lWidth
    l.height = lHeight
    l.Margin = 0
    l.ScaleFontSize = blnLabelTextScaleFontSize
    l.AutoSize = True
    l.FontSize = lngLabelFontSize
    l.LineStyle = lngLabelLineStyleNorm
    l.LineWidth = lngLabelLineWidthNorm
    l.ForeColour = lngLabelColor
    If l.LabelType = 0 Then
        l.Transparent = False
    Else
        If Button_mi3dCursor <> True Then
'            l.Outline = True
        End If
    End If
    Set GetNewLabel = l
End Function

Public Sub subDeleteAppointLabel(im As DicomImage, strL As String)
'------------------------------------------------
'功能：删除指定类型的标注
'参数：im--需要删除指定类型标注的图像；strL--需要删除的标注tag中包含的指定内容。
'返回：无，直接删除图像的指定标注
'2009用
'------------------------------------------------
    Dim i  As Integer
    If strL = "" Then Exit Sub
    For i = im.Labels.Count To G_INT_SYS_LABEL_COUNT Step -1
        If Mid(im.Labels(i).Tag, 1, Len(strL)) = strL Then im.Labels.Remove i
    Next
End Sub

Public Sub SubInitPeriod(im As DicomImage)
'------------------------------------------------
'功能：为每一幅图增加前n个系统句柄，n的数量由常量G_INT_SYS_LABEL_COUNT决定
'参数：im--需要增加系统标注（系统句柄）的图像。
'返回：无，直接在图像上增加n个系统句柄
'2009用
'------------------------------------------------
    Dim CurrentLabel As DicomLabel
    Dim i As Integer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If im.Labels.Count > 0 Then
        MsgBox "传入图像已经有对象，不能初始化句柄", vbInformation, gstrSysName
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To G_INT_SYS_LABEL_COUNT
        Set CurrentLabel = New DicomLabel
        CurrentLabel.LabelType = doLabelRectangle
        CurrentLabel.Transparent = False
        CurrentLabel.XOR = False
        CurrentLabel.BackColour = lngPeriodColor
        CurrentLabel.ForeColour = 0
        CurrentLabel.height = intPeriodSize
        CurrentLabel.width = CurrentLabel.height
        CurrentLabel.Visible = False
        CurrentLabel.left = G_INT_SYS_LABEL_HIDE_LEFT
        CurrentLabel.ImageTied = True
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If i > 1 And i < 6 Then  '''''作为裁减遮盖用的句柄
            CurrentLabel.BackColour = vbBlack
            CurrentLabel.ForeColour = vbBlack
        End If
        '''''''''''''''裁减用矩形框'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If i = 1 Then
            CurrentLabel.Transparent = True
            CurrentLabel.ForeColour = vbBlack 'vbBlue
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''体位标注专用句柄'''''''''''''''''
        If i >= G_INT_SYS_LABEL_TIWEI And i <= G_INT_SYS_LABEL_TIWEI + 3 Then
            CurrentLabel.LabelType = doLabelSpecial
            CurrentLabel.ScaleWithCell = True
            CurrentLabel.ScaleFontSize = blnpatientInfoScaleFontSize
            CurrentLabel.width = 200
            CurrentLabel.height = 200
'            CurrentLabel.BackColour = lngViewerBackColor
            CurrentLabel.ForeColour = lngpatientInfoColor
            CurrentLabel.Margin = 2
            CurrentLabel.ImageTied = False
            CurrentLabel.Transparent = True
        End If
        
        If i >= G_INT_SYS_LABEL_MPRV And i <= G_INT_SYS_LABEL_MPR_RESULT_V Then   '''矢冠状重建使用句柄
            If i = G_INT_SYS_LABEL_MPRV Or i = G_INT_SYS_LABEL_MPRH _
                Or i = G_INT_SYS_LABEL_MPR_RESULT_H Or i = G_INT_SYS_LABEL_MPR_RESULT_V Then    '两根控制线和结果图中的两个投影线
                CurrentLabel.LabelType = doLabelLine
                CurrentLabel.ForeColour = vbRed
                CurrentLabel.LineWidth = 2
            Else        '控制线上的五个控制点
                CurrentLabel.Transparent = False
                CurrentLabel.LabelType = doLabelEllipse
                CurrentLabel.LineWidth = 1
                CurrentLabel.ForeColour = RGB(255, 255, 255)
                CurrentLabel.width = G_INT_MPR_RADIUS
                CurrentLabel.height = G_INT_MPR_RADIUS
            End If
            CurrentLabel.ImageTied = True
        End If
        
        '''''''''''''30号标注,用于显示窗宽窗位'''''''''''''''''''''''''''''''''''''''''''''''''''''
        If i = G_INT_SYS_LABEL_WWWL Then
            CurrentLabel.LabelType = doLabelText
            CurrentLabel.width = 0          '宽度和高度的设置会影响AutoSize
            CurrentLabel.height = 0
            CurrentLabel.ImageTied = False  '此设置会影响ScaleWithCell
            CurrentLabel.Transparent = True
            CurrentLabel.ScaleWithCell = True
            CurrentLabel.ScaleFontSize = blnpatientInfoScaleFontSize
            CurrentLabel.AutoSize = True
'            CurrentLabel.BackColour = lngViewerBackColor
            CurrentLabel.ForeColour = lngpatientInfoColor
            CurrentLabel.left = 0
            CurrentLabel.Text = "WL"
            CurrentLabel.Visible = True
            CurrentLabel.Alignment = doAlignCentre
        End If
        
        '处理病人四角信息
        If i >= G_INT_SYS_LABEL_PAT_INFO And i <= G_INT_SYS_LABEL_PAT_INFO + 3 Then
            CurrentLabel.LabelType = doLabelText
            CurrentLabel.width = 0          '宽度和高度的设置会影响AutoSize
            CurrentLabel.height = 0
            CurrentLabel.ImageTied = False  '此设置会影响ScaleWithCell
            CurrentLabel.Transparent = True
            CurrentLabel.ScaleWithCell = True
            CurrentLabel.ScaleFontSize = blnpatientInfoScaleFontSize
            CurrentLabel.Font.Name = strPatientInfoFontName
            CurrentLabel.Font.Size = lngPatientInfoFontSize
            CurrentLabel.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
            CurrentLabel.Font.Italic = blnPatientInfoFontItalic
            CurrentLabel.AutoSize = True
            CurrentLabel.ForeColour = lngpatientInfoColor
            CurrentLabel.left = 0
            CurrentLabel.top = 0
            Select Case i
                Case G_INT_SYS_LABEL_PAT_INFO
                    CurrentLabel.Tag = "PAT1"
                    CurrentLabel.Alignment = doAlignLeft
                Case G_INT_SYS_LABEL_PAT_INFO + 1
                    CurrentLabel.Tag = "PAT2"
                    CurrentLabel.Alignment = doAlignBottomLeft
                Case G_INT_SYS_LABEL_PAT_INFO + 2
                    CurrentLabel.Tag = "PAT3"
                    CurrentLabel.Alignment = doAlignBottomRight
                Case G_INT_SYS_LABEL_PAT_INFO + 3
                    CurrentLabel.Tag = "PAT4"
                    CurrentLabel.Alignment = doAlignRight
            End Select
        End If
        
        '病人标尺信息和标尺单位
        If i >= G_INT_SYS_LABEL_RULLER And i <= G_INT_SYS_LABEL_RULLER + 7 Then
            If i >= G_INT_SYS_LABEL_RULLER + 4 Then
                CurrentLabel.LabelType = doLabelText    '标尺单位
                CurrentLabel.AutoSize = True
                CurrentLabel.width = 0          '宽度和高度的设置会影响AutoSize
                CurrentLabel.height = 0
                CurrentLabel.Transparent = True
                CurrentLabel.ScaleFontSize = blnpatientInfoScaleFontSize
                CurrentLabel.Font.Name = strPatientInfoFontName
                CurrentLabel.Font.Size = lngPatientInfoFontSize
                CurrentLabel.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
                CurrentLabel.Font.Italic = blnPatientInfoFontItalic
                CurrentLabel.AutoSize = True
                CurrentLabel.left = 0
                CurrentLabel.top = 0
            Else
                CurrentLabel.LabelType = doLabelRuler   '标尺
            End If
            CurrentLabel.ImageTied = False  '此设置会影响ScaleWithCell
            CurrentLabel.ScaleWithCell = True
            CurrentLabel.ScaleFontSize = False
            CurrentLabel.ForeColour = lngRulerLeftColor
            CurrentLabel.LineWidth = intRulerLineWidth
        End If
        
        '打印标记
        If i = G_INT_SYS_LABEL_PRINT_TAG Then
            CurrentLabel.LabelType = doLabelText
            CurrentLabel.width = 400
            CurrentLabel.height = lngPatientInfoFontSize * 4
            CurrentLabel.ImageTied = False  '此设置会影响ScaleWithCell,1000,1000
            CurrentLabel.Transparent = True
            CurrentLabel.ScaleWithCell = True
            CurrentLabel.ScaleFontSize = blnpatientInfoScaleFontSize
            CurrentLabel.Font.Name = strPatientInfoFontName
            CurrentLabel.Font.Size = lngPatientInfoFontSize
            CurrentLabel.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
            CurrentLabel.Font.Italic = blnPatientInfoFontItalic
            CurrentLabel.ForeColour = lngpatientInfoColor
            CurrentLabel.BackColour = vbRed
            CurrentLabel.left = 300
            CurrentLabel.top = 50
            CurrentLabel.Text = "已打印"
            CurrentLabel.ShowTextBox = True
            CurrentLabel.Shadow = doShadowBottomRight
            CurrentLabel.Alignment = doAlignCentre
        End If
        
        im.Labels.Add CurrentLabel
    Next
    im.Labels(1).TagObject = im.Labels(6)
    im.Labels(6).TagObject = im.Labels(1)
End Sub

Public Sub UpdateMarkers(Image As DicomImage, Optional blnShow As Boolean = True)
'------------------------------------------------
'功能：根据图像显示或隐藏病人体位信息
'参数：Image--需要显示体位信息的图像；blnShow--是否显示病人信息。
'返回：无，直接在图像上显示或隐藏体位信息。
'------------------------------------------------
    Dim DG As New DicomGlobal
    Dim l As DicomLabel, i As Integer
    DG.DirectionStrings = IIf(blnChinaMark, "右\左\前\后\脚\头", "R\L\A\P\I\S")
    If blnShow Then
        Set l = Image.Labels(G_INT_SYS_LABEL_TIWEI)
        If blnAnatomicMarkersLeft Then
            l.left = 0
            l.top = 500
            l.Text = "LEFT"
            l.Font.Name = strPatientInfoFontName
            l.Font.Size = lngPatientInfoFontSize
            l.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
            l.Font.Italic = blnPatientInfoFontItalic
            l.ScaleFontSize = blnpatientInfoScaleFontSize
            l.ForeColour = lngpatientInfoColor
            l.Visible = True
        Else
            l.Visible = False
            l.left = G_INT_SYS_LABEL_HIDE_LEFT
        End If
        
        Set l = Image.Labels(G_INT_SYS_LABEL_TIWEI + 1)
        If blnAnatomicMarkersTop Then
            l.left = 500 - l.width / 2
            l.top = 0
            l.Alignment = doAlignCentre
            l.Text = "TOP"
            l.Font.Name = strPatientInfoFontName
            l.Font.Size = lngPatientInfoFontSize
            l.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
            l.Font.Italic = blnPatientInfoFontItalic
            l.ScaleFontSize = blnpatientInfoScaleFontSize
            l.ForeColour = lngpatientInfoColor
            l.Visible = True
        Else
            l.Visible = False
            l.left = G_INT_SYS_LABEL_HIDE_LEFT
        End If
        
        Set l = Image.Labels(G_INT_SYS_LABEL_TIWEI + 2)
        If blnAnatomicMarkersRight Then
            l.left = 1000 - l.width
            l.top = 500
            l.Alignment = doAlignRight
            l.Text = "RIGHT"
            l.Font.Name = strPatientInfoFontName
            l.Font.Size = lngPatientInfoFontSize
            l.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
            l.Font.Italic = blnPatientInfoFontItalic
            l.ScaleFontSize = blnpatientInfoScaleFontSize
            l.ForeColour = lngpatientInfoColor
            l.Visible = True
        Else
            l.Visible = False
            l.left = G_INT_SYS_LABEL_HIDE_LEFT
        End If
        
        Set l = Image.Labels(G_INT_SYS_LABEL_TIWEI + 3)
        If blnAnatomicMarkersBottom Then
            l.left = 500 - l.width / 2
            l.top = 1000 - l.height
            l.Alignment = doAlignBottomCentre
            l.Text = "BOTTOM"
            l.Font.Name = strPatientInfoFontName
            l.Font.Size = lngPatientInfoFontSize
            l.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
            l.Font.Italic = blnPatientInfoFontItalic
            l.ScaleFontSize = blnpatientInfoScaleFontSize
            l.ForeColour = lngpatientInfoColor
            l.Visible = True
        Else
            l.Visible = False
            l.left = G_INT_SYS_LABEL_HIDE_LEFT
        End If
    Else        '隐藏体位标注
        For i = G_INT_SYS_LABEL_TIWEI To G_INT_SYS_LABEL_TIWEI + 3
            Set l = Image.Labels(i)
            l.Visible = False
            l.left = G_INT_SYS_LABEL_HIDE_LEFT
        Next i
    End If
    Image.Refresh True
End Sub


Public Function UpdateRuler(im As DicomImage, blnDisp As Boolean) As Long
'------------------------------------------------
'功能：显示图像标尺,直接显示或隐藏标尺
'参数：im--显示标尺的图像；blnDisp--是否显示标尺：True显示标尺；False不显示标尺
'返回： 0---正常；1--标尺标注数量不对；2-其他错误
'------------------------------------------------
    Dim l As DicomLabel
    Dim lUnit As DicomLabel
    
    On Error GoTo err
    
    '检查图像的标注标尺是否存在
    If im.Labels.Count < G_INT_SYS_LABEL_RULLER + 4 Then
        UpdateRuler = 1
        Exit Function
    End If
    
    Set l = im.Labels(G_INT_SYS_LABEL_RULLER)           '处理左标尺和标尺单位
    Set lUnit = im.Labels(G_INT_SYS_LABEL_RULLER + 4)
    l.left = IIf(blnDisp, intRulerLeft, G_INT_SYS_LABEL_HIDE_LEFT)
    l.top = intRulerTop
    l.width = intRulerWidth
    l.height = intRulerHeight
    l.ForeColour = lngRulerLeftColor
    l.LineWidth = intRulerLineWidth
    l.Visible = IIf(blnRulerDsipLeft And blnDisp, True, False)
    '病人信息使用的题头，0--不使用题头；1--中文题头；2--英文题头
    If lngPatientInfoTitle = 0 Then
        lUnit.Text = l.TickSpacing(False) & im.Labels(11).ROIDistanceUnits
    ElseIf lngPatientInfoTitle = 2 Then
        lUnit.Text = "Unit:" & l.TickSpacing(False) & im.Labels(11).ROIDistanceUnits
    Else
        lUnit.Text = "单位：" & l.TickSpacing(False) & im.Labels(11).ROIDistanceUnits
    End If
    lUnit.left = l.left
    lUnit.top = l.top + l.height
    lUnit.Font.Name = strPatientInfoFontName
    lUnit.Font.Size = lngPatientInfoFontSize
    lUnit.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    lUnit.Font.Italic = blnPatientInfoFontItalic
    lUnit.ScaleFontSize = blnpatientInfoScaleFontSize
    lUnit.ForeColour = l.ForeColour
    lUnit.LineWidth = l.LineWidth
    lUnit.Visible = l.Visible
    If im.Labels(11).ROIDistanceUnits = "Pixels" Then lUnit.Visible = False
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set l = im.Labels(G_INT_SYS_LABEL_RULLER + 1)               '处理上标尺和单位
    Set lUnit = im.Labels(G_INT_SYS_LABEL_RULLER + 5)
    l.left = IIf(blnDisp, intRulerTop, G_INT_SYS_LABEL_HIDE_LEFT)
    l.top = intRulerLeft
    l.width = intRulerHeight
    l.height = intRulerWidth
    l.ForeColour = lngRulerLeftColor
    l.LineWidth = intRulerLineWidth
    l.Visible = IIf(blnRulerDsipTop And blnDisp, True, False)
    lUnit.Text = "单位：" & l.TickSpacing(False) & im.Labels(11).ROIDistanceUnits
    lUnit.left = l.left + l.width
    lUnit.top = l.top
    lUnit.Font.Name = strPatientInfoFontName
    lUnit.Font.Size = lngPatientInfoFontSize
    lUnit.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    lUnit.Font.Italic = blnPatientInfoFontItalic
    lUnit.ScaleFontSize = blnpatientInfoScaleFontSize
    lUnit.ForeColour = l.ForeColour
    lUnit.LineWidth = l.LineWidth
    lUnit.Visible = l.Visible
    If im.Labels(11).ROIDistanceUnits = "Pixels" Then lUnit.Visible = False
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set l = im.Labels(G_INT_SYS_LABEL_RULLER + 2)           '处理右标尺和单位
    Set lUnit = im.Labels(G_INT_SYS_LABEL_RULLER + 6)
    l.left = IIf(blnDisp, 1000 - intRulerLeft, G_INT_SYS_LABEL_HIDE_LEFT)
    l.top = intRulerTop
    l.width = -intRulerWidth
    l.height = intRulerHeight
    l.ForeColour = lngRulerLeftColor
    l.LineWidth = intRulerLineWidth
    l.Visible = IIf(blnRulerDsipRight And blnDisp, True, False)
    lUnit.Text = "单位：" & l.TickSpacing(False) & im.Labels(11).ROIDistanceUnits
    lUnit.left = l.left + l.width
    lUnit.top = l.top + l.height
    lUnit.Font.Name = strPatientInfoFontName
    lUnit.Font.Size = lngPatientInfoFontSize
    lUnit.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    lUnit.Font.Italic = blnPatientInfoFontItalic
    lUnit.ScaleFontSize = blnpatientInfoScaleFontSize
    lUnit.ForeColour = l.ForeColour
    lUnit.LineWidth = l.LineWidth
    lUnit.Visible = l.Visible
    If im.Labels(11).ROIDistanceUnits = "Pixels" Then lUnit.Visible = False
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set l = im.Labels(G_INT_SYS_LABEL_RULLER + 3)           '处理下标尺和单位
    Set lUnit = im.Labels(G_INT_SYS_LABEL_RULLER + 7)
    l.left = IIf(blnDisp, intRulerTop, G_INT_SYS_LABEL_HIDE_LEFT)
    l.top = 1000 - intRulerLeft
    l.width = intRulerHeight
    l.height = -intRulerWidth
    l.ForeColour = lngRulerLeftColor
    l.LineWidth = intRulerLineWidth
    l.Visible = IIf(blnRulerDsipBottom And blnDisp, True, False)
    lUnit.Text = "单位：" & l.TickSpacing(False) & im.Labels(11).ROIDistanceUnits
    lUnit.left = l.left + l.width
    lUnit.top = l.top + l.height
    lUnit.Font.Name = strPatientInfoFontName
    lUnit.Font.Size = lngPatientInfoFontSize
    lUnit.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    lUnit.Font.Italic = blnPatientInfoFontItalic
    lUnit.ScaleFontSize = blnpatientInfoScaleFontSize
    lUnit.ForeColour = l.ForeColour
    lUnit.LineWidth = l.LineWidth
    lUnit.Visible = l.Visible
    If im.Labels(11).ROIDistanceUnits = "Pixels" Then lUnit.Visible = False
    Exit Function
    
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    UpdateRuler = 2
End Function


Public Sub subDispImageInfo(im As DicomImage, blnDisp As Boolean, blnRefreshPatiIn As Boolean, blnRefreshWL As Boolean, Optional strPatientInfo1 As String = "", Optional strPatientInfo2 As String = "", _
                     Optional strPatientInfo3 As String = "", Optional strPatientInfo4 As String = "")
'------------------------------------------------
'功能：显示或隐藏病人图像四角信息和窗宽窗位显示
'参数：im--显示病人信息的图像；blnDisp--显示或隐藏病人四角信息和窗宽窗位，True为显示，False为隐藏；
'      blnRefreshPatiIn--是否按照传入的四个四角信息字符串，刷新病人四角信息，True为刷新，False为不刷新；
'      blnRefreshWL -- 是否刷新图像的窗宽窗位
'      strPatientInfo1--左上角的病人信息；strPatientInfo2--左下角的病人信息；
'      strPatientInfo3--右下角的病人信息；strPatientInfo4--右上角的病人信息。
'返回：
'------------------------------------------------
    Dim i, j, intTop, intLeft As Integer
    Dim l As DicomLabel
    
    On Error GoTo err
    
    '左上
    Set l = im.Labels(G_INT_SYS_LABEL_PAT_INFO)
    l.Visible = blnDisp
    l.left = IIf(blnDisp, 0, G_INT_SYS_LABEL_HIDE_LEFT)
    l.ScaleFontSize = blnpatientInfoScaleFontSize
    l.ForeColour = lngpatientInfoColor
    l.Font.Name = strPatientInfoFontName
    l.Font.Size = lngPatientInfoFontSize
    l.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    l.Font.Italic = blnPatientInfoFontItalic
    l.Text = IIf(blnRefreshPatiIn, strPatientInfo1, l.Text)
    '左下
    Set l = im.Labels(G_INT_SYS_LABEL_PAT_INFO + 1)
    l.Visible = blnDisp
    l.left = IIf(blnDisp, 0, G_INT_SYS_LABEL_HIDE_LEFT)
    l.ScaleFontSize = blnpatientInfoScaleFontSize
    l.ForeColour = lngpatientInfoColor
    l.Font.Name = strPatientInfoFontName
    l.Font.Size = lngPatientInfoFontSize
    l.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    l.Font.Italic = blnPatientInfoFontItalic
    l.Text = IIf(blnRefreshPatiIn, strPatientInfo2, l.Text)
    '右下
    Set l = im.Labels(G_INT_SYS_LABEL_PAT_INFO + 2)
    l.Visible = blnDisp
    l.left = IIf(blnDisp, 0, G_INT_SYS_LABEL_HIDE_LEFT)
    l.ScaleFontSize = blnpatientInfoScaleFontSize
    l.Font.Name = strPatientInfoFontName
    l.Font.Size = lngPatientInfoFontSize
    l.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    l.Font.Italic = blnPatientInfoFontItalic
    l.ForeColour = lngpatientInfoColor
    l.Text = IIf(blnRefreshPatiIn, strPatientInfo3, l.Text)
    '右上
    Set l = im.Labels(G_INT_SYS_LABEL_PAT_INFO + 3)
    l.Visible = blnDisp
    l.left = IIf(blnDisp, 0, G_INT_SYS_LABEL_HIDE_LEFT)
    l.ScaleFontSize = blnpatientInfoScaleFontSize
    l.Font.Name = strPatientInfoFontName
    l.Font.Size = lngPatientInfoFontSize
    l.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    l.Font.Italic = blnPatientInfoFontItalic
    l.ForeColour = lngpatientInfoColor
    l.Text = IIf(blnRefreshPatiIn, strPatientInfo4, l.Text)
    ''''''窗宽窗位标注处理''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_WWWL).Visible = blnDisp
    im.Labels(G_INT_SYS_LABEL_WWWL).left = IIf(blnDisp, 0, G_INT_SYS_LABEL_HIDE_LEFT)
    im.Labels(G_INT_SYS_LABEL_WWWL).Font.Name = strPatientInfoFontName
    im.Labels(G_INT_SYS_LABEL_WWWL).Font.Size = lngPatientInfoFontSize
    im.Labels(G_INT_SYS_LABEL_WWWL).Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    im.Labels(G_INT_SYS_LABEL_WWWL).Font.Italic = blnPatientInfoFontItalic
    im.Labels(G_INT_SYS_LABEL_WWWL).ScaleFontSize = blnpatientInfoScaleFontSize
    im.Labels(G_INT_SYS_LABEL_WWWL).ForeColour = lngpatientInfoColor
    im.Labels(G_INT_SYS_LABEL_WWWL).Text = IIf(blnRefreshWL, "W:" & im.width & "-L:" & im.Level, im.Labels(G_INT_SYS_LABEL_WWWL).Text)
    If lngWinWidthLevelLocation = 1 Then  '''1-上边；2-下边；3-左边；4-右边
        im.Labels(G_INT_SYS_LABEL_WWWL).top = 0
        im.Labels(G_INT_SYS_LABEL_WWWL).Alignment = doAlignCentre
    ElseIf lngWinWidthLevelLocation = 2 Then
        im.Labels(G_INT_SYS_LABEL_WWWL).top = 0
        im.Labels(G_INT_SYS_LABEL_WWWL).Alignment = doAlignBottomCentre
    ElseIf lngWinWidthLevelLocation = 3 Then
        im.Labels(G_INT_SYS_LABEL_WWWL).top = 500
        im.Labels(G_INT_SYS_LABEL_WWWL).Alignment = doAlignLeft
    ElseIf lngWinWidthLevelLocation = 4 Then
        im.Labels(G_INT_SYS_LABEL_WWWL).top = 500
        im.Labels(G_INT_SYS_LABEL_WWWL).Alignment = doAlignRight
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub subGetImgInfoLabel(intSeriesIndex As Integer, intIndexType As Integer, img As DicomImage, strInfoLabel() As String, Optional lngPrefix As Long = 0, Optional blnIsOnlyExport As Boolean = False)
'------------------------------------------------
'功能：从图像中提取病人的四个角信息标注，配合系统参数设置中四个角标注的内容使用
'参数： intSeriesIndex -- 图像所在序列的索引
'       intIndexType -- 序列索引的类型，0--从ZLSeriesInfos提取，1 -- 从ZLShowSeriesInfos提取
'       img--提取病人信息的图像；
'       strInfoLabel()--放返回值的数组；
'       lngPrefix--表示使用前缀的类型。0-不用前缀；1-使用中文前缀；2-使用英文前缀
'返回：无，直接填写到strInfoLabel()数组里面。
'2009用
'------------------------------------------------
    Dim i As Integer, j As Integer
    Dim iLocation As Integer
    Dim v As Variant
    Dim iCount(4) As Integer
    Dim iMax As Integer
    Dim strInfo() As String
    Dim StrTmp As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If UBound(strInfoLabel) <> 4 Then
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To 4
        strInfoLabel(i) = ""
        iCount(i) = 0
    Next
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To lngInfoLabelCount
        If aInfoLabelLocate(i).bUsed Then
            iLocation = aInfoLabelLocate(i).lngLocation
            iCount(iLocation) = iCount(iLocation) + 1
        End If
    Next
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    iMax = iCount(1)
    If iMax < iCount(2) Then iMax = iCount(2)
    If iMax < iCount(3) Then iMax = iCount(3)
    If iMax < iCount(4) Then iMax = iCount(4)
            
    ReDim strInfo(4, iMax) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To lngInfoLabelCount
        If aInfoLabelLocate(i).bUsed And (Not blnIsOnlyExport Or aInfoLabelLocate(i).blnIsExport) Then
            iLocation = aInfoLabelLocate(i).lngLocation
            If (img.Attributes(Val("&H" & aInfoLabelLocate(i).strGroup), Val("&H" & aInfoLabelLocate(i).strElement)).Exists) _
                Or (aInfoLabelLocate(i).strGroup = "1" And aInfoLabelLocate(i).strElement = "1") _
                Or (aInfoLabelLocate(i).strGroup = "2" And aInfoLabelLocate(i).strElement = "2") _
                Or (aInfoLabelLocate(i).strGroup = "3" And aInfoLabelLocate(i).strElement = "3") _
                Or (aInfoLabelLocate(i).strGroup = "0010" And aInfoLabelLocate(i).strElement = "1010") Then
                
                If aInfoLabelLocate(i).strGroup = "1" And aInfoLabelLocate(i).strElement = "1" Then
                    '对于Tag为（1,1）的图像属性，需要按照其中文简称为标识，进行计算。
                    v = funcCalImgInfoLabel(img, aInfoLabelLocate(i).strCName)
                ElseIf aInfoLabelLocate(i).strGroup = "2" And aInfoLabelLocate(i).strElement = "2" Then
                    '对于Tag为（2,2)的图像属性，是用户定义的，直接显示中文名
                    v = aInfoLabelLocate(i).strCName
                ElseIf aInfoLabelLocate(i).strGroup = "3" And aInfoLabelLocate(i).strElement = "3" Then
                    '对于Tag为（3,3）的图像属性，是数据库字段，序列信息中提取预先存储的数据库信息
                    v = funGetDBInfoLabel(intSeriesIndex, aInfoLabelLocate(i).strCName, intIndexType)
                ElseIf aInfoLabelLocate(i).strGroup = "0020" And aInfoLabelLocate(i).strElement = "0013" Then
                    '图像号（20，13）特殊处理，如果是多帧图像，显示帧数
                    If img.FrameCount > 1 Then
                        v = img.Attributes(&H20, &H13).Value & "-" & img.Frame
                    Else
                        v = img.Attributes(&H20, &H13).Value
                    End If
                ElseIf aInfoLabelLocate(i).strGroup = "0010" And aInfoLabelLocate(i).strElement = "1010" Then
                    '年龄（0010,1010）特殊处理，如果年龄为空，则通过出生日期计算年龄
                    If Not img.Attributes(&H10, &H1010).Exists Or IsNull(img.Attributes(&H10, &H1010).Value) Then
                        If Not IsNull(img.DateOfBirth) Then
                            v = DateDiff("yyyy", img.DateOfBirth, Now)
                        Else
                            v = ""
                        End If
                    Else
                        v = img.Attributes(&H10, &H1010).Value
                    End If
                Else
                    v = img.Attributes(Val("&H" & aInfoLabelLocate(i).strGroup), Val("&H" & aInfoLabelLocate(i).strElement)).Value
                End If
                If TypeName(v) = "String()" Then
                    StrTmp = v(1)
                Else
                    StrTmp = IIf(IsNull(v), "", v)
                End If
                '将读出来的值插入到四个角的数组中
                If aInfoLabelLocate(i).strGroup = "2" And aInfoLabelLocate(i).strElement = "2" Then
                    strInfo(iLocation, aInfoLabelLocate(i).lngOrder) = StrTmp
                Else
                    If IsNull(v) Or StrTmp = "" Then
                        strInfo(iLocation, aInfoLabelLocate(i).lngOrder) = "--"
                    Else
                        Select Case lngPrefix
                        Case 0          ''不使用前缀
                            strInfo(iLocation, aInfoLabelLocate(i).lngOrder) = StrTmp
                        Case 1          ''使用中文前缀
                            strInfo(iLocation, aInfoLabelLocate(i).lngOrder) = aInfoLabelLocate(i).strCName & " " & StrTmp
                        Case 2          ''使用英文前缀
                            strInfo(iLocation, aInfoLabelLocate(i).lngOrder) = aInfoLabelLocate(i).strEName & " " & StrTmp
                        End Select
                    End If
                End If
            Else
                strInfo(iLocation, aInfoLabelLocate(i).lngOrder) = "--"
            End If
        End If
    Next
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '将文字组合起来，并且显示
    For i = 1 To 4
        For j = 0 To iCount(i) - 1
            If strInfo(i, j) <> "--" Then
                If strInfoLabel(i) = "" Then
                    strInfoLabel(i) = strInfo(i, j)
                Else
                    strInfoLabel(i) = strInfoLabel(i) & vbCrLf & strInfo(i, j)
                End If
            End If
        Next
    Next
End Sub


Public Sub subInitImageLabels(intSeriesIndex As Integer, intIndexType As Integer, img As DicomImage, blnShowLabel As Boolean, _
           Optional blnGetImgInfo As Boolean = False, Optional blnCreateSysLabel As Boolean = False, Optional blnIsOnlyExport As Boolean = False)
'------------------------------------------------
'功能：初始化、显示或隐藏指定图像的标注信息:系统标注；体位标注，标尺，四角信息，窗宽窗位。
'      只对一个图像进行操作。
'参数：img--需要处理图像标注信息的图像；blnShowLabel--显示或隐藏标注，True-显示标注，False-隐藏标注；
'      blnGetImgInfo-是否读取图像四角信息，True-从图像读取四角信息，False-不读取四角信息；
'      blnCreateSysLabel-是否创建系统标注，True-创建系统标注，False-不创建系统标注。
'返回：无，直接改变图像。
'2009用
'------------------------------------------------
    Dim strInfo(4) As String
    If blnCreateSysLabel Then SubInitPeriod img      ''初始化 G_INT_SYS_LABEL_COUNT 个系统句柄
    
    'If Not blnIsOnlyExport Then
        UpdateMarkers img, blnShowLabel   ''显示病人体位信息
        UpdateRuler img, blnShowLabel     ''显示病人标尺
    'End If
    
    If blnGetImgInfo Then
        subGetImgInfoLabel intSeriesIndex, intIndexType, img, strInfo, lngPatientInfoTitle, blnIsOnlyExport    ''从数据库读取病人四角信息
    End If
    
    subDispImageInfo img, blnShowLabel, blnGetImgInfo, blnGetImgInfo, strInfo(1), strInfo(2), strInfo(3), strInfo(4)   ''显示病人四角信息和窗宽窗位信息
End Sub


Private Function funGetDBInfoLabel(intSeriesIndex As Integer, strFieldName As String, intIndexType As Integer)
'------------------------------------------------
'功能：根据传入的中文简称，在序列信息中查找对应四角标注的显示值。
'参数： intSeriesIndex -- 图像所在序列的索引
'       strFieldName -- 需要显示的数据库信息的名称
'       intIndexType -- 序列索引的类型，0--从ZLSeriesInfos提取，1 -- 从ZLShowSeriesInfos提取
'返回：根据中文简称提取出来的显示值。
'------------------------------------------------
    funGetDBInfoLabel = Null
    
    On Error GoTo err
    
    If intIndexType = 0 Then
        If intSeriesIndex > 0 And intSeriesIndex <= ZLSeriesInfos.Count Then
            If strFieldName = "[姓名]" Then
                funGetDBInfoLabel = ZLSeriesInfos(intSeriesIndex).strCName
            ElseIf strFieldName = "[英文名]" Then
                funGetDBInfoLabel = ZLSeriesInfos(intSeriesIndex).strEName
            ElseIf strFieldName = "[性别]" Then
                funGetDBInfoLabel = ZLSeriesInfos(intSeriesIndex).strSex
            ElseIf strFieldName = "[年龄]" Then
                funGetDBInfoLabel = ZLSeriesInfos(intSeriesIndex).strAge
            ElseIf strFieldName = "[检查号]" Then
                funGetDBInfoLabel = ZLSeriesInfos(intSeriesIndex).strStudyID
            ElseIf strFieldName = "[医嘱ID]" Then
                funGetDBInfoLabel = ZLSeriesInfos(intSeriesIndex).strOrderID
            End If
        End If
    Else
        If intSeriesIndex > 0 And intSeriesIndex <= ZLShowSeriesInfos.Count Then
            If strFieldName = "[姓名]" Then
                funGetDBInfoLabel = ZLShowSeriesInfos(intSeriesIndex).strCName
            ElseIf strFieldName = "[英文名]" Then
                funGetDBInfoLabel = ZLShowSeriesInfos(intSeriesIndex).strEName
            ElseIf strFieldName = "[性别]" Then
                funGetDBInfoLabel = ZLShowSeriesInfos(intSeriesIndex).strSex
            ElseIf strFieldName = "[年龄]" Then
                funGetDBInfoLabel = ZLShowSeriesInfos(intSeriesIndex).strAge
            ElseIf strFieldName = "[检查号]" Then
                funGetDBInfoLabel = ZLShowSeriesInfos(intSeriesIndex).strStudyID
            ElseIf strFieldName = "[医嘱ID]" Then
                funGetDBInfoLabel = ZLShowSeriesInfos(intSeriesIndex).strOrderID
            End If
        End If
    End If
    Exit Function
err:
    '不做任何处理，返回空
    funGetDBInfoLabel = Null
End Function

Private Function funcCalImgInfoLabel(img As DicomImage, strCalName As String) As Variant
'------------------------------------------------
'功能：根据传入的中文简称，来计算对应四角标注的显示值。
'参数：img--需要显示四角标注的图像；aInfo--标注信息类型。
'返回：根据中文简称计算出来的显示值。
'2009用
'------------------------------------------------
    funcCalImgInfoLabel = Null
    Dim v1 As Variant
    Dim v2 As Variant
    Dim v3 As Variant
    If strCalName = "层厚层距" Then
        If img.Attributes(&H18, &H50).Exists Then
            v1 = img.Attributes(&H18, &H50).Value   'slice thickness
            v2 = Null
            If IsNull(v1) Then Exit Function
            If img.Attributes(&H18, &H88).Exists Then
                v2 = img.Attributes(&H18, &H88).Value   'spacing between slices
            End If
            If IsNull(v2) Then
                funcCalImgInfoLabel = v1 & "thk"
            Else
                funcCalImgInfoLabel = v1 & "thk/" & v2 - v1 & "sp"
            End If
        End If
    ElseIf strCalName = "视野FOV" Then
        If img.Attributes(&H28, &H10).Exists And img.Attributes(&H28, &H11).Exists And img.Attributes(&H28, &H30).Exists Then
            v1 = img.Attributes(&H28, &H10).Value   'rows
            v2 = img.Attributes(&H28, &H11).Value   'columns
            v3 = img.Attributes(&H28, &H30).Value  'pixel spacing
            If IsNull(v1) Or IsNull(v2) Or IsNull(v3) Then
                Exit Function
            End If
            '针对贵阳肺科医院，北京国药恒瑞美联信息技术公司的DR特殊处理，它的像素距离字段只有一维值
            If TypeName(v3) = "String()" Then
                If UBound(v3) < 2 Then
                    Exit Function
                End If
            End If
            funcCalImgInfoLabel = Format(v1 * v3(1) / 10, "#00.0") & " CM X " & Format(v2 * v3(2) / 10, "#00.0") & " CM"
        End If
    End If
End Function


Public Sub subSaveLabelToImg(img As DicomImage)
'------------------------------------------------
'功能：将标注保存到DICOM图像的头信息里面
'参数：img--需要保存标注的图像。
'返回：无，直接将标注信息填写到图像的头信息里面。
'2009用
'------------------------------------------------
    Dim la As DicomLabel
    Dim ds As DicomDataSet
    Dim dssAll As DicomDataSets
    Dim i As Integer
    Dim iIncrease As Integer
    Dim lngTemp As Long
    Dim strPoints As String
    Dim j As Integer
    Dim vPoints As Variant
    Dim lngPointsCount As Long
    Dim aSaveTagObject() As Integer
    ReDim aSaveTagObject(img.Labels.Count) As Integer
    
    Dim v As Variant
    '图像中标注全部都是保留标注，前几十个标注是系统保留的标注
    If img.Labels.Count <= G_INT_SYS_LABEL_COUNT Then
        Exit Sub
    End If
    Set dssAll = New DicomDataSets
    iIncrease = 0
    For i = G_INT_SYS_LABEL_COUNT + 1 To img.Labels.Count
        '保存标注到img中
        Set la = img.Labels(i)
        Set ds = New DicomDataSet
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("TINDEX").Group) + iIncrease)), Val("&h" & cLabelStore("TINDEX").Element), cLabelStore("TINDEX").VR, iIncrease
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Left").Group) + iIncrease)), Val("&h" & cLabelStore("Left").Element), cLabelStore("Left").VR, la.left
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Top").Group) + iIncrease)), Val("&h" & cLabelStore("Top").Element), cLabelStore("Top").VR, la.top
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Width").Group) + iIncrease)), Val("&h" & cLabelStore("Width").Element), cLabelStore("Width").VR, la.width
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Height").Group) + iIncrease)), Val("&h" & cLabelStore("Height").Element), cLabelStore("Height").VR, la.height
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("LabelType").Group) + iIncrease)), Val("&h" & cLabelStore("LabelType").Element), cLabelStore("LabelType").VR, la.LabelType
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("ImageTied").Group) + iIncrease)), Val("&h" & cLabelStore("ImageTied").Element), cLabelStore("ImageTied").VR, la.ImageTied
        
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Alignment").Group) + iIncrease)), Val("&h" & cLabelStore("Alignment").Element), cLabelStore("Alignment").VR, la.Alignment
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("AnchorImageTied").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorImageTied").Element), cLabelStore("AnchorImageTied").VR, la.AnchorImageTied
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("AnchorX").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorX").Element), cLabelStore("AnchorX").VR, la.AnchorX
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("AnchorY").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorY").Element), cLabelStore("AnchorY").VR, la.AnchorY
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Angle").Group) + iIncrease)), Val("&h" & cLabelStore("Angle").Element), cLabelStore("Angle").VR, la.Angle
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("AutoSize").Group) + iIncrease)), Val("&h" & cLabelStore("AutoSize").Element), cLabelStore("AutoSize").VR, la.AutoSize
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("BackColour").Group) + iIncrease)), Val("&h" & cLabelStore("BackColour").Element), cLabelStore("BackColour").VR, la.BackColour
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("BackStyle").Group) + iIncrease)), Val("&h" & cLabelStore("BackStyle").Element), cLabelStore("BackStyle").VR, la.BackStyle
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("FontName").Group) + iIncrease)), Val("&h" & cLabelStore("FontName").Element), cLabelStore("FontName").VR, la.FontName
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("FontSize").Group) + iIncrease)), Val("&h" & cLabelStore("FontSize").Element), cLabelStore("FontSize").VR, la.FontSize
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("ForeColour").Group) + iIncrease)), Val("&h" & cLabelStore("ForeColour").Element), cLabelStore("ForeColour").VR, la.ForeColour
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("LineStyle").Group) + iIncrease)), Val("&h" & cLabelStore("LineStyle").Element), cLabelStore("LineStyle").VR, la.LineStyle
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("LineWidth").Group) + iIncrease)), Val("&h" & cLabelStore("LineWidth").Element), cLabelStore("LineWidth").VR, la.LineWidth
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Margin").Group) + iIncrease)), Val("&h" & cLabelStore("Margin").Element), cLabelStore("Margin").VR, la.Margin
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Outline").Group) + iIncrease)), Val("&h" & cLabelStore("Outline").Element), cLabelStore("Outline").VR, la.Outline
        
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("RotateTextWithImage").Group) + iIncrease)), Val("&h" & cLabelStore("RotateTextWithImage").Element), cLabelStore("RotateTextWithImage").VR, la.RotateTextWithImage
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("ScaleFontSize").Group) + iIncrease)), Val("&h" & cLabelStore("ScaleFontSize").Element), cLabelStore("ScaleFontSize").VR, la.ScaleFontSize
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("ScaleWithCell").Group) + iIncrease)), Val("&h" & cLabelStore("ScaleWithCell").Element), cLabelStore("ScaleWithCell").VR, la.ScaleWithCell
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Shadow").Group) + iIncrease)), Val("&h" & cLabelStore("Shadow").Element), cLabelStore("Shadow").VR, la.Shadow
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("ShowAnchor").Group) + iIncrease)), Val("&h" & cLabelStore("ShowAnchor").Element), cLabelStore("ShowAnchor").VR, la.ShowAnchor
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("ShowTextBox").Group) + iIncrease)), Val("&h" & cLabelStore("ShowTextBox").Element), cLabelStore("ShowTextBox").VR, la.ShowTextBox
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Tag").Group) + iIncrease)), Val("&h" & cLabelStore("Tag").Element), cLabelStore("Tag").VR, la.Tag
        
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Text").Group) + iIncrease)), Val("&h" & cLabelStore("Text").Element), cLabelStore("Text").VR, la.Text
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Transparent").Group) + iIncrease)), Val("&h" & cLabelStore("Transparent").Element), cLabelStore("Transparent").VR, la.Transparent
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Visible").Group) + iIncrease)), Val("&h" & cLabelStore("Visible").Element), cLabelStore("Visible").VR, la.Visible
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("XOR").Group) + iIncrease)), Val("&h" & cLabelStore("XOR").Element), cLabelStore("XOR").VR, la.XOR
        
        '需要特殊处理的类型
        'Points类型
        lngPointsCount = UBound(la.Points)
        If lngPointsCount = 0 Then
            ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Points").Group) + iIncrease)), Val("&h" & cLabelStore("Points").Element), cLabelStore("Points").VR, 0
            ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("PointsCount").Group) + iIncrease)), Val("&h" & cLabelStore("PointsCount").Element), cLabelStore("PointsCount").VR, lngPointsCount
        Else
            vPoints = la.Points
            strPoints = vPoints(1)
            For j = 2 To lngPointsCount
                strPoints = strPoints & ";" & vPoints(j)
            Next
            ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Points").Group) + iIncrease)), Val("&h" & cLabelStore("Points").Element), cLabelStore("Points").VR, strPoints
            ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("PointsCount").Group) + iIncrease)), Val("&h" & cLabelStore("PointsCount").Element), cLabelStore("PointsCount").VR, lngPointsCount
        End If
        
        'TagObject类型
        If la.TagObject Is Nothing Then
            ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("TagObject").Group) + iIncrease)), Val("&h" & cLabelStore("TagObject").Element), cLabelStore("TagObject").VR, 0
        Else
            lngTemp = img.Labels.IndexOf(la.TagObject)
            aSaveTagObject(i) = iIncrease + 1
            ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("TagObject").Group) + iIncrease)), Val("&h" & cLabelStore("TagObject").Element), cLabelStore("TagObject").VR, lngTemp
        End If
        
        '将一个标注添加到数据集中
        dssAll.Add ds
        iIncrease = iIncrease + 1
    Next
    '处理特殊的标注属性
    'TagObject类型
    For i = 1 To dssAll.Count
            Set ds = dssAll(i)
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("TagObject").Group) + i - 1)), Val("&h" & cLabelStore("TagObject").Element)).Value
            lngTemp = v(1)
            ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("TagObject").Group) + i - 1)), Val("&h" & cLabelStore("TagObject").Element), cLabelStore("TagObject").VR, aSaveTagObject(lngTemp)
    Next
    If dssAll.Count > 0 Then    '往图像里面添加标注
        img.Attributes.AddExplicit Val("&h" & cLabelStore("TPRODUCER").Group), Val("&h" & cLabelStore("TPRODUCER").Element), cLabelStore("TPRODUCER").VR, cProducer
        img.Attributes.AddExplicit Val("&h" & cLabelStore("TSUM").Group), Val("&h" & cLabelStore("TSUM").Element), cLabelStore("TSUM").VR, iIncrease
        img.Attributes.AddExplicit Val("&h" & cLabelStore("TALL").Group), Val("&h" & cLabelStore("TALL").Element), cLabelStore("TALL").VR, dssAll
    Else                        '本图像没有标注，将原有标注清空
        img.Attributes.Remove Val("&h" & cLabelStore("TPRODUCER").Group), Val("&h" & cLabelStore("TPRODUCER").Element)
        img.Attributes.Remove Val("&h" & cLabelStore("TSUM").Group), Val("&h" & cLabelStore("TSUM").Element)
        img.Attributes.Remove Val("&h" & cLabelStore("TALL").Group), Val("&h" & cLabelStore("TALL").Element)
    End If
End Sub


Public Sub subReadLabelFromImg(img As DicomImage)
'------------------------------------------------
'功能：从图像的头文件中读取标注，并显示标注
'参数：img--需要读取标注的图像。
'返回：无，直接将图像中的标注读取出来，并进行显示。
'2009用
'------------------------------------------------
    Dim ds As DicomDataSet
    Dim dss As DicomDataSets
    Dim las As New DicomLabels
    Dim la As DicomLabel
    Dim v As Variant
    Dim i As Integer
    Dim iCount As Integer
    Dim lngTemp As Long
    Dim aTagObject() As Long
    Dim iIncrease As Integer
    Dim strPoints As String
    Dim lngPointsCount As Long
    Dim aReadTagObject() As Long
    Dim j As Long
    Dim strX As String
    Dim strY As String
    Dim iOldCount As Integer
    ReDim aReadTagObject(cLabelStore.Count) As Long
    aReadTagObject(0) = 0
    
    If (img.Attributes(Val("&h" & cLabelStore("TPRODUCER").Group), Val("&h" & cLabelStore("TPRODUCER").Element)).Exists) Then
        v = img.Attributes(Val("&h" & cLabelStore("TPRODUCER").Group), Val("&h" & cLabelStore("TPRODUCER").Element)).Value
        If IsNull(v) Or v <> cProducer Then
            Exit Sub
        End If
        
        If IsNull(img.Attributes(Val("&h" & cLabelStore("TALL").Group), Val("&h" & cLabelStore("TALL").Element))) Then
            Exit Sub
        End If
        Set dss = img.Attributes(Val("&h" & cLabelStore("TALL").Group), Val("&h" & cLabelStore("TALL").Element)).Value
        iCount = dss.Count
        iIncrease = 0
        For i = 1 To iCount
            Set ds = dss(i)
            Set la = New DicomLabel
                        
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Left").Group) + iIncrease)), Val("&h" & cLabelStore("Left").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Left").Group) + iIncrease)), Val("&h" & cLabelStore("Left").Element))
            la.left = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Top").Group) + iIncrease)), Val("&h" & cLabelStore("Top").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Top").Group) + iIncrease)), Val("&h" & cLabelStore("Top").Element))
            la.top = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Width").Group) + iIncrease)), Val("&h" & cLabelStore("Width").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Width").Group) + iIncrease)), Val("&h" & cLabelStore("Width").Element))
            la.width = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Height").Group) + iIncrease)), Val("&h" & cLabelStore("Height").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Height").Group) + iIncrease)), Val("&h" & cLabelStore("Height").Element))
            la.height = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("LabelType").Group) + iIncrease)), Val("&h" & cLabelStore("LabelType").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("LabelType").Group) + iIncrease)), Val("&h" & cLabelStore("LabelType").Element))
            la.LabelType = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ImageTied").Group) + iIncrease)), Val("&h" & cLabelStore("ImageTied").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ImageTied").Group) + iIncrease)), Val("&h" & cLabelStore("ImageTied").Element))
            la.ImageTied = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Alignment").Group) + iIncrease)), Val("&h" & cLabelStore("Alignment").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Alignment").Group) + iIncrease)), Val("&h" & cLabelStore("Alignment").Element))
            la.Alignment = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("AnchorImageTied").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorImageTied").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("AnchorImageTied").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorImageTied").Element))
            la.AnchorImageTied = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("AnchorX").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorX").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("AnchorX").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorX").Element))
            la.AnchorX = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("AnchorY").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorY").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("AnchorY").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorY").Element))
            la.AnchorY = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Angle").Group) + iIncrease)), Val("&h" & cLabelStore("Angle").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Angle").Group) + iIncrease)), Val("&h" & cLabelStore("Angle").Element))
            la.Angle = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("AutoSize").Group) + iIncrease)), Val("&h" & cLabelStore("AutoSize").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("AutoSize").Group) + iIncrease)), Val("&h" & cLabelStore("AutoSize").Element))
            la.AutoSize = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("BackColour").Group) + iIncrease)), Val("&h" & cLabelStore("BackColour").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("BackColour").Group) + iIncrease)), Val("&h" & cLabelStore("BackColour").Element))
            la.BackColour = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("BackStyle").Group) + iIncrease)), Val("&h" & cLabelStore("BackStyle").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("BackStyle").Group) + iIncrease)), Val("&h" & cLabelStore("BackStyle").Element))
            la.BackStyle = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("FontName").Group) + iIncrease)), Val("&h" & cLabelStore("FontName").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("FontName").Group) + iIncrease)), Val("&h" & cLabelStore("FontName").Element))
            la.FontName = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("FontSize").Group) + iIncrease)), Val("&h" & cLabelStore("FontSize").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("FontSize").Group) + iIncrease)), Val("&h" & cLabelStore("FontSize").Element))
            la.FontSize = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ForeColour").Group) + iIncrease)), Val("&h" & cLabelStore("ForeColour").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ForeColour").Group) + iIncrease)), Val("&h" & cLabelStore("ForeColour").Element))
            la.ForeColour = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("LineStyle").Group) + iIncrease)), Val("&h" & cLabelStore("LineStyle").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("LineStyle").Group) + iIncrease)), Val("&h" & cLabelStore("LineStyle").Element))
            la.LineStyle = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("LineWidth").Group) + iIncrease)), Val("&h" & cLabelStore("LineWidth").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("LineWidth").Group) + iIncrease)), Val("&h" & cLabelStore("LineWidth").Element))
            la.LineWidth = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Margin").Group) + iIncrease)), Val("&h" & cLabelStore("Margin").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Margin").Group) + iIncrease)), Val("&h" & cLabelStore("Margin").Element))
            la.Margin = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Outline").Group) + iIncrease)), Val("&h" & cLabelStore("Outline").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Outline").Group) + iIncrease)), Val("&h" & cLabelStore("Outline").Element))
            la.Outline = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("RotateTextWithImage").Group) + iIncrease)), Val("&h" & cLabelStore("RotateTextWithImage").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("RotateTextWithImage").Group) + iIncrease)), Val("&h" & cLabelStore("RotateTextWithImage").Element))
            la.RotateTextWithImage = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ScaleFontSize").Group) + iIncrease)), Val("&h" & cLabelStore("ScaleFontSize").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ScaleFontSize").Group) + iIncrease)), Val("&h" & cLabelStore("ScaleFontSize").Element))
            la.ScaleFontSize = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ScaleWithCell").Group) + iIncrease)), Val("&h" & cLabelStore("ScaleWithCell").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ScaleWithCell").Group) + iIncrease)), Val("&h" & cLabelStore("ScaleWithCell").Element))
            la.ScaleWithCell = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Shadow").Group) + iIncrease)), Val("&h" & cLabelStore("Shadow").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Shadow").Group) + iIncrease)), Val("&h" & cLabelStore("Shadow").Element))
            la.Shadow = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ShowAnchor").Group) + iIncrease)), Val("&h" & cLabelStore("ShowAnchor").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ShowAnchor").Group) + iIncrease)), Val("&h" & cLabelStore("ShowAnchor").Element))
            la.ShowAnchor = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ShowTextBox").Group) + iIncrease)), Val("&h" & cLabelStore("ShowTextBox").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ShowTextBox").Group) + iIncrease)), Val("&h" & cLabelStore("ShowTextBox").Element))
            la.ShowTextBox = v(1)
            
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Tag").Group) + iIncrease)), Val("&h" & cLabelStore("Tag").Element))
            If IsNull(v) Then
                la.Tag = ""
            Else
                la.Tag = v(1)
            End If
            
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Text").Group) + iIncrease)), Val("&h" & cLabelStore("Text").Element))
            If IsNull(v) Then
                la.Text = ""
            Else
                la.Text = v(1)
            End If
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Transparent").Group) + iIncrease)), Val("&h" & cLabelStore("Transparent").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Transparent").Group) + iIncrease)), Val("&h" & cLabelStore("Transparent").Element))
            la.Transparent = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Visible").Group) + iIncrease)), Val("&h" & cLabelStore("Visible").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Visible").Group) + iIncrease)), Val("&h" & cLabelStore("Visible").Element))
            la.Visible = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("XOR").Group) + iIncrease)), Val("&h" & cLabelStore("XOR").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("XOR").Group) + iIncrease)), Val("&h" & cLabelStore("XOR").Element))
            la.XOR = v(1)
            
            '需要特殊处理的类型
            'Points类型
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Points").Group) + iIncrease)), Val("&h" & cLabelStore("Points").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Points").Group) + iIncrease)), Val("&h" & cLabelStore("Points").Element))
            
            strPoints = v(1)
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("PointsCount").Group) + iIncrease)), Val("&h" & cLabelStore("PointsCount").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("PointsCount").Group) + iIncrease)), Val("&h" & cLabelStore("PointsCount").Element))
            lngPointsCount = v(1) / 2
            For j = 1 To lngPointsCount - 1
                strX = left(strPoints, InStr(strPoints, ";") - 1)
                strPoints = Right(strPoints, Len(strPoints) - InStr(strPoints, ";"))
                strY = left(strPoints, InStr(strPoints, ";") - 1)
                strPoints = Right(strPoints, Len(strPoints) - InStr(strPoints, ";"))
                la.AddPoint Val(strX), Val(strY)
            Next
            If lngPointsCount > 0 Then
                strX = left(strPoints, InStr(strPoints, ";") - 1)
                strPoints = Right(strPoints, Len(strPoints) - InStr(strPoints, ";"))
                strY = strPoints
                la.AddPoint Val(strX), Val(strY)
            End If
            
            las.Add la
            iIncrease = iIncrease + 1
        Next
    End If
    '将标注放到图像里面
    '先将图像原来的标注数量记下来
    iOldCount = img.Labels.Count
    For i = 1 To las.Count
        img.Labels.Add las(i)
    Next
    
    '处理特殊的属性类型
    '处理TagObject
    For i = 1 To las.Count
        Set ds = dss(i)
        Set la = img.Labels(iOldCount + i)
        v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("TagObject").Group) + i - 1)), Val("&h" & cLabelStore("TagObject").Element))
        
        lngTemp = v(1)
        If lngTemp = 0 Then
            Set la.TagObject = Nothing
        Else
            la.TagObject = img.Labels(iOldCount + lngTemp)
        End If
    Next
End Sub

Public Sub subDrawRefLine(imgSource As DicomImage, imgDest As DicomImage, blnCheckSpacing As Boolean, _
    strLineTag As String, blnShowNum As Boolean)
'------------------------------------------------
'功能：画定位线
'参数： imgSource--定位线的投影图
'       imgDest -- 定位线所在的图像
'       blnCheckSpacing -- 是否检测定位线之间的距离
'       strLineTag -- 定位线的Tag的内容
'       blnShowNum -- 是否显示数字
'返回：无
'2009用
'------------------------------------------------
    Dim l As DicomLabel
    Dim dlNum As DicomLabel
    Dim iXoffset As Integer, iYoffset As Integer
    Dim strIOPSource As String
    Dim strIOPDest As String

    '（0020,0052）判断Frame of Reference UID是否相同，只能对参考帧UID相同的图像做定位线操作
    If Not IsNull(imgDest.Attributes(&H20, &H52).Value) And Not IsNull(imgSource.Attributes(&H20, &H52).Value) Then
        If imgDest.Attributes(&H20, &H52).Value = imgSource.Attributes(&H20, &H52).Value Then
            
            '对同一个层面的图像不做定位线
            If imgSource.Attributes(&H20, &H37).VM = 6 And imgDest.Attributes(&H20, &H37).VM = 6 Then
                strIOPSource = CInt(imgSource.Attributes(&H20, &H37).ValueByIndex(1)) & "," _
                                & CInt(imgSource.Attributes(&H20, &H37).ValueByIndex(2)) & "," _
                                & CInt(imgSource.Attributes(&H20, &H37).ValueByIndex(3)) & "," _
                                & CInt(imgSource.Attributes(&H20, &H37).ValueByIndex(4)) & "," _
                                & CInt(imgSource.Attributes(&H20, &H37).ValueByIndex(5)) & "," _
                                & CInt(imgSource.Attributes(&H20, &H37).ValueByIndex(6))
                strIOPDest = CInt(imgDest.Attributes(&H20, &H37).ValueByIndex(1)) & "," _
                                & CInt(imgDest.Attributes(&H20, &H37).ValueByIndex(2)) & "," _
                                & CInt(imgDest.Attributes(&H20, &H37).ValueByIndex(3)) & "," _
                                & CInt(imgDest.Attributes(&H20, &H37).ValueByIndex(4)) & "," _
                                & CInt(imgDest.Attributes(&H20, &H37).ValueByIndex(5)) & "," _
                                & CInt(imgDest.Attributes(&H20, &H37).ValueByIndex(6))
                If strIOPSource <> strIOPDest Then
                    Set l = imgDest.ReferenceLine(imgSource, True)
                    If l.LabelType = 3 Then
                        
                        If blnCheckSpacing = True Then
                            '判断当前定位线和上一条定位线之间的距离是否小于定位线间距，如果小于，则不显示
                            If imgDest.Labels(imgDest.Labels.Count - 1).Tag <> "RLL" _
                                Or Abs(imgDest.Labels(imgDest.Labels.Count - 1).left - l.left) >= lngReferenceLineSpacing _
                                Or Abs(imgDest.Labels(imgDest.Labels.Count - 1).top - l.top) >= lngReferenceLineSpacing Then
                                '可以画定位线，则不退出
                            Else
                                Exit Sub
                            End If
                        End If
                        
                        l.ForeColour = lngReferenceLineColor
                        l.Tag = strLineTag
                        l.LineStyle = lngReferenceLineStyle
                        imgDest.Labels.Add l
                        
                        If blnShowNum Then
                            Set dlNum = New DicomLabel
                            If Abs(l.width) > Abs(l.height) Then
                                iXoffset = 10
                                iYoffset = 0
                            Else
                                iXoffset = 0
                                iYoffset = 20
                            End If
                            dlNum.left = IIf(l.width > 0, l.left - iXoffset, l.left + iXoffset)
                            If dlNum.left < 0 Then
                                dlNum.left = 0
                            ElseIf dlNum.left > imgDest.sizex Then
                                dlNum.left = imgDest.sizex
                            End If
                            dlNum.top = IIf(l.height > 0, l.top - iYoffset, l.top + iYoffset)
                            If dlNum.top < 0 Then
                                dlNum.top = 0
                            ElseIf dlNum.top > imgDest.sizey Then
                                dlNum.top = imgDest.sizey
                            End If
                            
                            dlNum.LabelType = doLabelText
                            dlNum.Tag = strLineTag
                            dlNum.ForeColour = lngReferenceLineColor
                            dlNum.Text = IIf(Not IsNull(imgSource.Attributes(&H20, &H13).Value), imgSource.Attributes(&H20, &H13).Value, "")
                            dlNum.ImageTied = True
                            dlNum.FontSize = 12
                            imgDest.Labels.Add dlNum
                        End If
                    End If
                
                End If
            End If
        End If
    End If
End Sub


Public Function funDrawVas(lblLine As DicomLabel, img As DicomImage, intVasType As Integer) As Boolean
'------------------------------------------------
'功能：根据lblLine做自动血管测量
'参数：lblLine--进行血管测量的血管垂直线；img--进行血管测量的图像；intVasType--血管测量类型：1为正常血管，2为狭窄血管。
'返回：无
'2009用
'------------------------------------------------
    '计算血管壁
    Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long
    Dim lngRadius As Long
    Dim lngArea As Long
    Dim lblVas1 As DicomLabel
    Dim lblVas2 As DicomLabel
    Dim lblText As DicomLabel
    
    If lblLine.TagObject Is Nothing Or lblLine.TagObject.TagObject Is Nothing _
        Or lblLine.TagObject.TagObject.TagObject Is Nothing Then
       Exit Function
    End If
    Set lblText = lblLine.TagObject
    Set lblVas1 = lblText.TagObject
    Set lblVas2 = lblVas1.TagObject
    If funGetVasEdge(img, lblLine, IIf(intVasType = 1, intStandardThreshold, intNarrowThreshold), x1, y1, x2, y2) = True Then
        '设置血管壁短直线
        
        subDrawVasEdgeLine lblLine, lblVas1, x1, y1
        lblVas1.Text = x1 & "," & y1
        subDrawVasEdgeLine lblLine, lblVas2, x2, y2
        lblVas2.Text = x2 & "," & y2
        lngRadius = Sqr((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
        lngArea = 3.14 * lngRadius * lngRadius / 4
        lblText.Text = IIf(intVasType = 1, "正常血管直径：", "狭窄血管直径：") & lngRadius & _
                        "(" & lblLine.ROIDistanceUnits & ")" & vbCrLf & "血管面积：" _
                        & lngArea & "(sq " & lblLine.ROIDistanceUnits & ")"
        lblLine.Text = lngRadius & ":" & IIf(intVasType = 1, intStandardThreshold, intNarrowThreshold)
        funDrawVas = True
    End If
End Function
                       
Public Sub subChangeLabelForPrint(img As DicomImage, intType As Integer)
'------------------------------------------------
'功能：修改图像中四角标注、体位标注、窗宽窗位标注成跟图像一起缩放，为胶片打印做准备
'参数：img－－需要修改标注的图像,intType -- 0用于显示；1用于打印
'返回：无
'------------------------------------------------
    Dim dlLabel As DicomLabel
    Dim i As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strImageType As String
    Dim strImageFontSize As Integer
    Dim strImageAutoZoom As Boolean
    Dim strPostureFontSize As Integer
    Dim strPostureAutoZoom As Boolean
    Dim blnFontInverse As Boolean
    Dim blnFontShadow As Boolean
    Dim blnFontTransparent As Boolean
    
    strImageType = IIf(IsNull(img.Attributes(&H8, &H60).Value), "OT", img.Attributes(&H8, &H60).Value)
    
    If blLocalRun = True Then
        strSQL = "select 影像类别,字体大小,是否随图像缩放,体位标注字体大小,体位标注随图像缩放 from 影像胶片打印字体 where 影像类别 = '" & strImageType & "'"
        Set rsTmp = cnAccess.Execute(strSQL)
    Else
        strSQL = "select 影像类别,字体大小,是否随图像缩放,体位标注字体大小,体位标注随图像缩放,字体反色,字体阴影,字体背景透明 from 影像胶片打印字体 where 影像类别 = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strImageType)
    End If
    
    If rsTmp.EOF = True Then
        '设置默认字体：四角信息标注和普通测量标注
        Select Case strImageType
            Case "CR"
                strImageFontSize = 14
                strImageAutoZoom = True
            Case "CT"
                strImageFontSize = 20
                strImageAutoZoom = True
            Case "MR"
                strImageFontSize = 8
                strImageAutoZoom = False
            Case "RF"
                strImageFontSize = 18
                strImageAutoZoom = False
            Case Else
                strImageFontSize = 14
                strImageAutoZoom = True
        End Select
        
        '设置默认字体：体位标注
        If img.sizex > 1024 Then
            strPostureFontSize = 40
        ElseIf img.sizex > 512 Then
            strPostureFontSize = 25
        ElseIf img.sizex > 400 Then
            strPostureFontSize = 18
        Else
            strPostureFontSize = 10
        End If
        strPostureAutoZoom = True
        
        blnFontInverse = False
        blnFontShadow = False
        blnFontTransparent = True
    Else
        strImageFontSize = NVL(rsTmp("字体大小"), 14)
        strImageAutoZoom = NVL(rsTmp("是否随图像缩放"), "True")
        strPostureFontSize = NVL(rsTmp("体位标注字体大小"), 25)
        strPostureAutoZoom = NVL(rsTmp("体位标注随图像缩放"), "True")
        blnFontInverse = NVL(rsTmp("字体反色"), "False")
        blnFontShadow = NVL(rsTmp("字体阴影"), "False")
        blnFontTransparent = NVL(rsTmp("字体背景透明"), "True")
    End If
    
    
    For Each dlLabel In img.Labels
        '对以下标注，需要设置其文字大小
        '1、类型为doLabelSpecial的四边体位标注；
        '2、Tag =“PAT”为病人四角信息；
        '3、窗宽窗位标注
        '4、标尺
        '5、用户自己画的标注，且不包括体位标注
        
        If dlLabel.LabelType = doLabelSpecial Or Mid(dlLabel.Tag, 1, 3) = "PAT" Or _
            img.Labels.IndexOf(dlLabel) = G_INT_SYS_LABEL_WWWL Or _
            (img.Labels.IndexOf(dlLabel) >= G_INT_SYS_LABEL_RULLER And img.Labels.IndexOf(dlLabel) <= G_INT_SYS_LABEL_RULLER + 5) Or _
            img.Labels.IndexOf(dlLabel) > G_INT_SYS_LABEL_COUNT Then
            
            If intType = 0 Then     '用于显示
                dlLabel.ScaleFontSize = True
                dlLabel.ForeColour = vbWhite
                dlLabel.Shadow = doShadowNone
                dlLabel.Transparent = True
                dlLabel.XOR = False
            ElseIf intType = 1 Then     '用于打印
                If InStr(dlLabel.Tag, POSTURE_LABEL) = 0 Then
                    '设置四角标注和普通测量标注字体大小
                    dlLabel.FontSize = strImageFontSize
                    dlLabel.ScaleFontSize = strImageAutoZoom
                    
                Else
                    '设置体位标注字体大小
                    dlLabel.FontSize = strPostureFontSize
                    dlLabel.ScaleFontSize = strPostureAutoZoom
                End If
                If blnFontShadow = True Then
                    dlLabel.Shadow = doShadowTopLeft
                End If
                If blnFontTransparent = False Then
                    dlLabel.BackColour = vbBlack
                    dlLabel.Transparent = False
                End If
                If blnFontInverse = True Then
                    dlLabel.XOR = True
                End If
            End If
        Else
            If intType = 0 Then
                dlLabel.ScaleFontSize = True
            End If
        End If
    Next
    
    '隐藏当前显示的标注选择句柄
    For i = 11 To 20
        img.Labels(i).Visible = False
    Next i
        
    '隐藏打印标记
    img.Labels(G_INT_SYS_LABEL_PRINT_TAG).Visible = False
    
    '如果图像不是CT，则隐藏窗宽窗位
    If UCase(strImageType) = "CT" Then
        img.Labels(G_INT_SYS_LABEL_WWWL).Visible = True
    Else
        img.Labels(G_INT_SYS_LABEL_WWWL).Visible = False
    End If
End Sub

Public Sub funcGetCadioThoracicRatio(thisLabel As DicomLabel, thisImage As DicomImage)
'计算并且返回心胸比，返回值的格式是"0.xx"
'参数： thisLabel---心胸比的测量标注
'       thisImage---计算心胸比的图像
'心胸比测量的标注是“CTR1L”+“CTR1T”+“CTR2L”+“CTR2T”，四个标注连接起来的。
'thisLabel指向“CTR1L”或者“CTR2L”
    
    If thisLabel Is Nothing Then Exit Sub
    If thisLabel.TagObject Is Nothing Then Exit Sub
    If thisLabel.TagObject.TagObject Is Nothing Then Exit Sub
    If thisLabel.TagObject.TagObject.TagObject Is Nothing Then Exit Sub
    
    Dim intLine1 As Integer
    Dim intLine2 As Integer
    Dim otherLabel As DicomLabel
    
    On Error GoTo err
        
    If thisImage.RotateState = doRotateLeft Or thisImage.RotateState = doRotateRight Then
        intLine1 = Abs(thisLabel.height)
        intLine2 = Abs(thisLabel.TagObject.TagObject.height)
    Else
        intLine1 = Abs(thisLabel.width)
        intLine2 = Abs(thisLabel.TagObject.TagObject.width)
    End If
        
    If intLine1 = 0 Or intLine2 = 0 Then Exit Sub
        
    Set otherLabel = thisLabel.TagObject.TagObject
    If thisLabel.Tag = "CTR1L" Then     'intLine1是心脏线
        thisLabel.TagObject.Text = funROIResultString(thisLabel, thisImage)
        otherLabel.TagObject.Text = funROIResultString(otherLabel, thisImage) & vbCrLf _
            & "心胸比： " & Format(intLine1 / intLine2, "0.00")
    ElseIf thisLabel.Tag = "CTR2L" Then 'intLine1是胸廓线
        thisLabel.TagObject.Text = funROIResultString(thisLabel, thisImage) & vbCrLf _
            & "心胸比： " & Format(intLine2 / intLine1, "0.00")
        otherLabel.TagObject.Text = funROIResultString(otherLabel, thisImage)
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Sub ShowOverlay(f As frmViewer)
'------------------------------------------------
'功能：显示或者隐藏Overlay信息
'参数： f - 观片主窗体
'返回：无
'------------------------------------------------
    Dim v As DicomViewer
    Dim img As DicomImage
    
    On Error GoTo err
    
    For Each v In f.Viewer
        If v.Index <> 0 Then
            For Each img In v.Images
                If img.Attributes(&H6000, &H10).Exists = True Then
                    img.OverlayVisible(0) = Button_miShowOverlay
                End If
            Next
        End If
        v.Refresh
    Next
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
