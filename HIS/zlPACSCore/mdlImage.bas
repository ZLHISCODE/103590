Attribute VB_Name = "mdlImage"
Option Explicit
'--------------------------------------------------------
'功  能：本模块为图像处理的函数过程等
'编制人：胡涛，黄捷
'编制日期：2004.6.12
'过程函数清单：
'    subMPRLinenPhase():        对指定Viewer中所有图像的矢冠状重建控制点和控制线，按照指定图像中的设置进行同步
'    subInitMPRLine()：         对当前窗体中被选中Viewer中全部图像，初始化矢冠状重建控制点和控制线
'    funcPlaneRestructInit():   矢冠状重建初始化，同时填写图像的层厚、总高度和像素数组
'    LeagelToACRebuild():       判断图像是否满足矢冠状重建的条件
'    subGetArray():             根据Line1控制线，在Viewer的第一个图像上的位置，计算出该直线中每一个点的像素坐标数组
'    subACRebuild():            对灰度值数组进行插值和平滑处理
'    subGetLabelStoreToVar():   从数据库中读取标注保存所使用的TAG到系统变量
'    subaCorrectCursor():       鼠标移动如果超出图像范围则修正其鼠标位置
'    funAutoWinWL():            自适应调窗
'    subStackEnd():             穿梭结束
'    subLabelCopyRebuild():     重建图像的标注关联关系
'    ResizeRegion():            自动计算指定区域内，一定数目图像可排列的行列数目
'    subSetWidthLevelF():       设置窗宽窗位功能键弹出菜单
'    GetAngle():                计算通过三个点连成的两条线之间的角度
'    Max7InArray():             从数组里面取值最大的7个下标，对其求平均值
'    funIsShutter():            判断输入的影像类别是否需要进行图像消隐操作
'    subDrawImgShutter():       根据系统设置的影像类别，给输入的图像画图像消隐
'    funGetLinePoints():        从图像的给定线型标注（直线、折线）中提取灰度值数组和起点、终点坐标
'    funGetVasEdge():           用于血管狭窄测量，根据直线标注和预设的阈值，查找血管壁的坐标。
'    subDrawVasEdgeLine():      用于血管狭窄测量，根据直线标注和血管壁的坐标，确定并画出血管壁短直线的位置。
'    subCenterZoom()：          对图像进行缩放。以当前viewer中心点为缩放中心点。
'修改记录：
'    2005.7.07    黄捷
'    2005.8.19    黄捷
'    2005.9.15    黄捷
'    2006-2-10    黄捷
'-------------------------------------------------------

Public ToltalHeight As Integer                         ''重建的总高度
Public aPixels() As Integer                                  ''保存重建像素值的数组

Public Sub subMPRLinenPhase(v As DicomViewer, im As DicomImage)
'------------------------------------------------
'功能：对指定Viewer中所有图像的矢冠状重建控制点和控制线，按照指定图像中的设置进行同步
'参数：v--进行图像中矢冠状控制标注同步的Viewer；im--做为同步标准的图像
'返回：无，直接将v中所有图像的矢冠状控制点线进行设置。
'2009 要修改，改成只对显示的图像进行修改
'------------------------------------------------
    Dim img As DicomImage, i As Integer
    For Each img In v.Images
        For i = G_INT_SYS_LABEL_MPRV To G_INT_SYS_LABEL_MPR_POINT_O
            img.Labels(i).Visible = im.Labels(i).Visible
            img.Labels(i).left = im.Labels(i).left
            img.Labels(i).width = im.Labels(i).width
            img.Labels(i).top = im.Labels(i).top
            img.Labels(i).height = im.Labels(i).height
        Next
        img.Refresh False
    Next
End Sub

Public Sub subInitMPRLine(thisViewer As DicomViewer)
'------------------------------------------------
'功能：对当前窗体中被选中Viewer中全部图像，初始化矢冠状重建控制点和控制线
'参数：     thisViewer--进行矢冠状重建的Viewer
'返回：无，直接对im图像上的矢冠状重建标注做初始化。
'2009用
'------------------------------------------------
     Dim im As DicomImage
     
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
     If thisViewer.Images.Count = 0 Then Exit Sub
     
     Set im = thisViewer.Images(1)
     
     Call funInitMPRControlLines(im, True)
     
     subMPRLinenPhase thisViewer, im
End Sub

Public Function funInitMPRControlLines(im As DicomImage, blnVisible As Boolean)
'------------------------------------------------
'功能：初始化指定图像的矢冠状重建控制点和控制线
'参数：     im--进行矢冠状重建的轴位图像
'返回：无，直接对im图像上的矢冠状重建标注做初始化。
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
    
    For i = G_INT_SYS_LABEL_MPRV To G_INT_SYS_LABEL_MPR_POINT_O
        im.Labels(i).Visible = blnVisible
        If i >= G_INT_SYS_LABEL_MPR_POINT_V1 Then
            im.Labels(i).width = G_INT_MPR_RADIUS
            im.Labels(i).height = G_INT_MPR_RADIUS
        End If
    Next
    
    im.Labels(G_INT_SYS_LABEL_MPRV).left = im.sizex / 2
    im.Labels(G_INT_SYS_LABEL_MPRV).top = 0
    im.Labels(G_INT_SYS_LABEL_MPRV).height = im.sizey
    im.Labels(G_INT_SYS_LABEL_MPRV).width = 0
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPRH).left = 0
    im.Labels(G_INT_SYS_LABEL_MPRH).top = im.sizey / 2
    im.Labels(G_INT_SYS_LABEL_MPRH).height = 0
    im.Labels(G_INT_SYS_LABEL_MPRH).width = im.sizex
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_V1).left = im.sizex / 2 - G_INT_MPR_RADIUS / 2
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_V1).top = -G_INT_MPR_RADIUS / 2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_V2).left = im.sizex / 2 - G_INT_MPR_RADIUS / 2
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_V2).top = im.sizey - G_INT_MPR_RADIUS / 2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_H1).left = -G_INT_MPR_RADIUS / 2
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_H1).top = im.sizey / 2 - G_INT_MPR_RADIUS / 2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_H2).left = im.sizex - G_INT_MPR_RADIUS / 2
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_H2).top = im.sizey / 2 - G_INT_MPR_RADIUS / 2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left = im.sizex / 2 - G_INT_MPR_RADIUS / 2
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top = im.sizey / 2 - G_INT_MPR_RADIUS / 2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Refresh False
     
    funInitMPRControlLines = True
    Exit Function
err:
    If ErrCenter = 1 Then Resume
End Function

Public Function funcPlaneRestructInit(Viewer As DicomViewer, thisForm As frmViewer) As Boolean
'------------------------------------------------
'功能：矢冠状重建初始化，同时填写图像的层厚、总高度和像素数组
'参数： viewer--进行矢冠状重建的viewer
'       thisForm -- 显示图像的窗体
'返回：True--初始化成功可以进行重建；False--初始化失败，不能够进行重建
'2009用
'------------------------------------------------
    Dim iHeight As Integer
    Dim iPixSpacing As Double
    Dim v As Variant
    Dim i As Integer
    Dim ix As Integer
    Dim iy As Integer
    
    funcPlaneRestructInit = False
    
    On Error GoTo err
    ''''获取图像的层厚'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''在前面已经对以下的 层厚，位置，图像数量，像素距离作了检查，本函数中不用再检查了。
    v = Viewer.Images(1).Attributes(&H28, &H30).Value
    iPixSpacing = v(1)
    iHeight = Viewer.Images.Count
    ''''''确定图像的总高度，在两个Slice Location的差，和层厚相叠的结果之间取最大值'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ToltalHeight = Abs(Viewer.Images(Viewer.Images.Count).Attributes(&H20, &H1041).Value - Viewer.Images(1).Attributes(&H20, &H1041).Value) / iPixSpacing
    
    '''''''''''''''''''重定义和填充像素值数组''''''''''''''''''''''''''''''''''''''
    zl9ComLib.zlCommFun.ShowFlash "正在初始化MPR重建，请等待！", thisForm
    zl9ComLib.zlCommFun.ShowFlash
    
    '如果定义MPR的三维数组超出内存许可范围，会出现“内存溢出”错误，aPixels维度=0，后续就直接用图像数据做MPR
    ReDim aPixels(Viewer.Images(1).sizex, Viewer.Images(1).sizey, Viewer.Images.Count) As Integer
    For i = 1 To Viewer.Images.Count
        v = Viewer.Images(i).Pixels
        For ix = 1 To Viewer.Images(i).sizex
            For iy = 1 To Viewer.Images(i).sizey
                aPixels(ix, iy, i) = v(ix, iy, 1)
            Next
        Next
    Next
    funcPlaneRestructInit = True
    zl9ComLib.zlCommFun.StopFlash
    Exit Function
err:
    funcPlaneRestructInit = False
    zl9ComLib.zlCommFun.StopFlash
End Function

Private Function LeagelToMPR(imgs As DicomImages) As Long
'------------------------------------------------
'功能：判断图像是否满足矢冠状重建的条件
'参数：imgs进行矢冠状重建的图像集
'返回：0--可以进行重建；1--不能进行重建

'------------------------------------------------
    Dim i As Long
    Dim j As Long
    
    LeagelToMPR = 1
    
    ''''''图像数量是否够3幅'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (imgs.Count < 3) Then
       Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim SeriesUID As String
    Dim thickness As Double
    Dim location() As Double
    ReDim location(imgs.Count) As Double
    Dim v As Variant
    Dim PixelSpacing As Double
    '''''''保存第一个图像的序列UID''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not IsNull(imgs(1).Attributes(&H20, &HE)) Then
        SeriesUID = imgs(1).Attributes(&H20, &HE).Value
    Else
        Exit Function
    End If
    '''''''保存第一个图像的层厚''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not IsNull(imgs(1).Attributes(&H18, &H50)) Then
        thickness = imgs(1).Attributes(&H18, &H50).Value
    Else
        Exit Function
    End If
    '''''''''保存第一个图像的切片位置Slice Location''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not IsNull(imgs(1).Attributes(&H20, &H1041)) Then
           location(1) = imgs(1).Attributes(&H20, &H1041).Value
    Else
       Exit Function
    End If
    ''''''''保存第一个图像的像素间距'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not IsNull(imgs(1).Attributes(&H28, &H30)) Then
        v = imgs(1).Attributes(&H28, &H30).Value
        PixelSpacing = v(1)
    Else
        Exit Function
    End If
    '''''''对其他图像做循环，判断是否满足条件''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 2 To imgs.Count
        '''''判断是否有相同的序列UID''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Not IsNull(imgs(i).Attributes(&H20, &HE)) Then
            If SeriesUID <> imgs(i).Attributes(&H20, &HE).Value Then
                Exit Function
            End If
        Else
            Exit Function
        End If
        ''''''''判断是否有相同的层厚'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Not IsNull(imgs(i).Attributes(&H18, &H50)) Then
            If thickness <> imgs(i).Attributes(&H18, &H50).Value Then
                Exit Function
            End If
        Else
            Exit Function
        End If
        ''''''''保存图像的位置SliceLocation'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Not IsNull(imgs(i).Attributes(&H20, &H1041)) Then
           location(i) = imgs(i).Attributes(&H20, &H1041).Value
        Else
           Exit Function
        End If
        '''''''判断是否有相同的像素间距''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Not IsNull(imgs(i).Attributes(&H28, &H30)) Then
            v = imgs(i).Attributes(&H28, &H30).Value
            If PixelSpacing <> v(1) Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    Next
    ''''''''判断是否有不相同的切片位置'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To imgs.Count
       For j = 1 To imgs.Count - i
          If location(i) = location(i + j) Then
             Exit Function
          End If
       Next
    Next
    '''''满足条件则返回真''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    LeagelToMPR = 0
End Function

Public Function LeagelToACRebuild(imgs As DicomImages) As Long
'------------------------------------------------
'功能：判断图像是否满足矢冠状重建的条件
'参数：imgs进行矢冠状重建的图像集
'返回：0--可以进行重建；1--不能进行重建，有提示
'------------------------------------------------
    
    LeagelToACRebuild = LeagelToMPR(imgs)
    
    If LeagelToACRebuild = 1 Then
        MsgBox "图像不能进行矢冠状重建，不满足以下条件之一：" & vbCrLf & vbCrLf & _
              "3幅图像以上；同一序列；相同层厚；不同位置；相同像素距离。", vbInformation, gstrSysName
    End If
End Function

Public Sub subGetArray(Line1 As DicomLabel, Image As DicomImage, LineLong() As POINTAPI)
'------------------------------------------------
'功能：根据Line1控制线，在Viewer的第一个图像上的位置，计算出该直线中每一个点的像素坐标数组
'参数：Line1--矢冠状控制线；Image--进行矢冠状重建的图像；LineLong()--做为返回值用，保存直线上所有点的坐标。
'返回：无，直接将直线中每一个点的坐标数组放到LineLong（）中。
'2009用
'------------------------------------------------
    Dim beginx As Integer, beginy As Integer
    Dim endx As Integer, endy As Integer
    Dim i As Integer
    Dim j As Integer
    Dim iH As Long
    Dim iW As Long
    Dim sizex As Integer
    Dim sizey As Integer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sizex = Image.sizex
    sizey = Image.sizey
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Line1.width > 0 Then
        beginx = Line1.left
        beginy = Line1.top
        endx = Line1.left + Line1.width
        endy = Line1.top + Line1.height
    Else
        endx = Line1.left
        endy = Line1.top
        beginx = Line1.left + Line1.width
        beginy = Line1.top + Line1.height
    End If
    ''''修正''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If beginx <= 0 Then beginx = 1
    If beginx > sizex Then beginx = sizex
    If beginy <= 0 Then beginy = 1
    If beginy > sizey Then beginy = sizey
    If endx <= 0 Then endx = 1
    If endx > sizex Then endx = sizex
    If endy <= 0 Then endy = 1
    If endy > sizey Then endy = sizey
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    iH = Abs(endy - beginy) + 1
    iW = Abs(endx - beginx) + 1
    ''''''''''''宽大于高''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If iW > iH Then
        ReDim LineLong(iW)
        j = 1
        For i = beginx To endx
            LineLong(j).x = i
            LineLong(j).y = IIf(((endy - beginy) / (endx - beginx)) * (i - beginx) + beginy > sizey, _
                    sizey, ((endy - beginy) / (endx - beginx)) * (i - beginx) + beginy)
            j = j + 1
        Next
     '''''''''高度大于宽度''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Else
        ReDim LineLong(iH)
        j = 1
        If beginy > endy Then
        ''''''交换begin和end的位置'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim iTmep As Integer
            iTmep = beginx
            beginx = endx
            endx = iTmep
            iTmep = beginy
            beginy = endy
            endy = iTmep
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = beginy To endy
            LineLong(j).y = i
            LineLong(j).x = IIf(((endx - beginx) / (endy - beginy)) * (i - beginy) + beginx > sizex, _
                  sizex, ((endx - beginx) / (endy - beginy)) * (i - beginy) + beginx)
            j = j + 1
        Next
    End If
End Sub

Public Sub subACRebuild(a() As Integer, b() As Integer)
'------------------------------------------------
'功能：对灰度值数组进行插值和平滑处理
'参数：a()--保存图像原来灰度值的二维数组，第一维是行，第二维是列；
'      b()--保存图像重建后新灰度值的二维数组，第一维是行，第二维是列；
'返回：无，直接使用b()来保存重建生成的新灰度数组。
'2009用
'------------------------------------------------

    '''''''''''''a为保存图像灰度值的二维数组，第一维是行，第二维是列'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim TolHeight As Long
   Dim lngWidth As Long
   Dim lngHeight As Long
   Dim dblThick As Double
   Dim dblResidua As Double
   Dim dblAccResidua As Double
   Dim lngThick As Long
   Dim lngThickAddOne As Long
   Dim intRealRows As Integer
   
   On Error GoTo err
   
   lngWidth = UBound(a, 1)      '图像灰度值数组第一维的长度，图像中控制线的长度
   lngHeight = UBound(a, 2)     '图像灰度值数组第二维的长度，图像的数量
   TolHeight = UBound(b, 2)
   
    ''''''''''从a中读取一行，到b中变换出来行数为SThickness指定的数量,对a内每一行做循环''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim i As Long, j As Long, k As Long
   dblThick = TolHeight / lngHeight
   lngThick = Int(dblThick)             '去尾取整，不做四舍五入
   lngThickAddOne = lngThick + 1
   dblResidua = dblThick - lngThick     '取余数
   dblAccResidua = 0
   dblAccResidua = dblAccResidua + dblResidua     '累加余数
   intRealRows = 0
   For i = 0 To lngHeight - 1
        dblAccResidua = dblAccResidua + dblResidua
        If dblAccResidua >= 1 Then
            dblAccResidua = dblAccResidua - 1
            For j = 1 To lngWidth
                For k = 1 To lngThickAddOne
                    b(j, intRealRows + k) = a(j, i + 1)
                Next
            Next
            intRealRows = intRealRows + lngThickAddOne
        Else
            For j = 1 To lngWidth
                For k = 1 To lngThick
                    b(j, intRealRows + k) = a(j, i + 1)
                Next
            Next
            intRealRows = intRealRows + lngThick
        End If
        
   Next
    ''''''对数组b中的点做平滑处理,采用加权模板，做两次平滑''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call funImageSmoothing(b, Int(IIf(dblThick / 2 > 5, 5, dblThick / 2)))
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub subGetLabelStoreToVar()
'------------------------------------------------
'功能：从数据库中读取标注保存所使用的TAG到系统变量
'参数：无
'返回：无
'2009用
'------------------------------------------------
   Dim strSQL As String
   Dim cOneAttr As clsLabelAttr
   Dim i As Integer
   '''''''获取系统变量的总数''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   If blLocalRun = True Then
      strSQL = "SELECT VGroup,Element,VR,标注属性 FROM 影像标注存储表"
      Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
   Else
      strSQL = "SELECT VGroup,Element,VR,标注属性 FROM 影像标注存储表"
      Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
   End If
   '''''清空集合里面的内容''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   For i = 1 To cLabelStore.Count
       cLabelStore.Remove 1
   Next i
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   On Error GoTo err
   
   With rsTemp
       .MoveFirst
       While Not .EOF
           Set cOneAttr = New clsLabelAttr
           cOneAttr.AttrName = !标注属性
           cOneAttr.Group = !VGroup
           cOneAttr.Element = !Element
           cOneAttr.VR = !VR
           ''''''加到集合里面'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           cLabelStore.Add cOneAttr, cOneAttr.AttrName
           .MoveNext
       Wend
   End With
   Exit Sub
err:
   If blLocalRun = True Then
        MsgBox "发生错误:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Public Sub subaCorrectCursor(v As DicomViewer, im As DicomImage, xx As Long, Yy As Long)
'------------------------------------------------
'功能：鼠标移动如果超出图像范围则修正其鼠标位置
'参数：v--图像所在的viewer；im--鼠标所在的图像；xx--鼠标所在的x方向位置，如果鼠标超出图像则将此值修改到图像之内；
'      yy--鼠标所在的y方向位置，如果鼠标超出图像则将此值修改到图像之内；
'返回：无
'2009用
'------------------------------------------------
    Dim x As Integer, y As Integer, w As Long, h As Long
    Dim i As DicomImage
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    w = v.width / v.MultiColumns / Screen.TwipsPerPixelX - v.CellSpacing * 2
    h = v.height / v.MultiRows / Screen.TwipsPerPixelY - v.CellSpacing * 2
    x = im.OriginX + v.CellSpacing
    y = im.OriginY + v.CellSpacing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If xx < x Then xx = x
    If xx > x + w Then xx = x + w
    If Yy < y Then Yy = y
    If Yy > y + h Then Yy = y + h
End Sub

Function funAutoWinWL(img As DicomImage, left As Long, top As Long, width As Long, _
                      height As Long, ByRef ww As Long, ByRef wl As Long) As Boolean
'------------------------------------------------
'功能：自适应调窗。
'算法说明：采用的方法是从给定的矩形区域中，提取全部像素点的灰度值，窗宽为该区域内最大灰度值和最小灰度值之间差的90%。
'         窗位为当前灰度值最多的7个点的灰度平均值。
'参数：img--需要进行自适应调窗的图像；(Left,Top,Width ,Height)--在图像上需要进行自适应调窗的矩形区域；
'      ww--返回窗宽值；wl--返回窗位值。
'返回：True--执行成功；Fasle--执行失败。
'2009用
'------------------------------------------------
    Dim iBitWidth As Integer
    Dim tImg As DicomImage
    Dim lngMax As Long
    Dim lngMin As Long
    Dim aImg As Variant
    ''''''初始化返回值''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    funAutoWinWL = False
    ww = img.width
    wl = img.Level
    '''''''''获取图像的存储位数信息,若此信息不存在，则返回错误'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not IsNull(img.Attributes(&H28, &H100)) Then
        iBitWidth = img.Attributes(&H28, &H100).Value
        iBitWidth = iBitWidth / 8
    Else
        Exit Function
    End If
    ''''''''对于宽度和高度同时为1的图像区域，不计算其自适应窗宽窗位，因为此时通过子图得不到最大和最小像素值''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Abs(width) <= 1 And Abs(height) <= 1 Then
        Exit Function
    End If
    ''''''''判断图像区域是否原图像''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (left = 0) And (top = 0) And (width = img.sizex) And (height = img.sizey) Then
        Set tImg = img
    Else
        ''''''''根据输入计算选取矩形区域的左上角和高宽，此时高宽需要是正数''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If width < 0 Then
            left = left + width
            width = -width
        End If
        If height < 0 Then
            top = top + height
            height = -height
        End If
        Set tImg = img.SubImage(left, top, width, height, 1, 1)
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If tImg.MinimumPixelValue(False) > 0 Then
        aImg = tImg.Histogram(tImg.MinimumPixelValue(False), tImg.MaximumPixelValue(False), iBitWidth)
        wl = tImg.MinimumPixelValue(False) + Max7InArray(aImg, lngMax, lngMin) * iBitWidth
    End If
    ww = Abs(tImg.MaximumPixelValue(False) - tImg.MinimumPixelValue(False)) * 0.9
    funAutoWinWL = True
End Function

Public Sub subStackEnd(v As DicomViewer, f As frmViewer)
'------------------------------------------------
'功能：穿梭结束
'参数：v--进行穿梭的viewer；f--穿梭的窗体。
'返回：无
'2009用
'------------------------------------------------
    Dim i As Integer
    i = f.MSFViewer.TextMatrix(f.intSelectedSerial, 3)
    v.Images.Add f.objStackOldImage
    v.Images.Move v.Images.Count, i
    subLabelCopyRebuild f.objStackOldImage, v.Images(i)
    v.Images.Remove i + 1
End Sub

Public Sub subLabelCopyRebuild(Simg As DicomImage, oImg As DicomImage)
'------------------------------------------------
'功能：重建图像的标注关联关系
'参数：sImg--源图像；oImg--目标图像
'返回：无
'2009用
'------------------------------------------------
    Dim l As DicomLabel
    For Each l In oImg.Labels
        If Not l.TagObject Is Nothing Then
            If Simg.Labels.IndexOf(l.TagObject) <> 0 Then
                Set l.TagObject = oImg.Labels(Simg.Labels.IndexOf(l.TagObject))
            End If
        End If
    Next
End Sub

Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer, _
    Optional ByVal MaxRows As Integer = 0, Optional ByVal MaxCols As Integer = 0)
'-----------------------------------------------------------------------------
'功能：根据输入的图像数量，图像区域的宽度和高度，返回最佳的图像排列行数和列数
'参数： ImageCount－－图像数量
'       RegionWidth--图像显示区域的宽度
'       RegionHeight--图像显示区域的高度
'       Rows－－[返回]最佳行数
'       Cols－－[返回]最佳列数
'       MaxRows －－可选，最大行数
'       MaxCols －－可选，最大列数
'返回：返回最佳行数Rows，最佳列数Cols
'2009用
'-----------------------------------------------------------------------------
    Dim iCols As Integer, iRows As Integer
    
    On Error GoTo err
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))

    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    Do While iRows * iCols > ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols - 1
        Else
            iRows = iRows - 1
        End If
    Loop
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    
    '根据最大值最终确定行数和列数
    If (MaxRows <> 0) And (MaxRows < iRows) Then
        Rows = MaxRows
    Else
        Rows = iRows
    End If
    If (MaxCols <> 0) And (MaxCols < iCols) Then
        Cols = MaxCols
    Else
        Cols = iCols
    End If
'不做错误处理
err:
End Sub

Public Sub subSetFilterF(im As DicomImage, f As frmViewer, Optional cbrPopup As CommandBarPopup)
'------------------------------------------------
'功能：设置滤镜模板功能键弹出菜单
'参数： im--设置滤镜模板的基准图像，提取图像的Modality
'       f--设置弹出菜单的窗体；
'       cbrPopup -- 在这个菜单项中增加弹出菜单
'返回：无
'-----------------------------------------------
    Dim strModality As String
    Dim ControlPopup As CommandBarPopup
    Dim cbrToolBar As CommandBarControl
    Dim i As Integer
    Dim MenuPopup As CommandBarPopup    '主菜单中的弹出菜单项
    Dim cbrMenuBar As CommandBarControl '主菜单中的菜单项
    
    If im Is Nothing Then Exit Sub
    If IsNull(im.Attributes(&H8, &H60).Value) Then Exit Sub         '获取Modality
    strModality = UCase(im.Attributes(&H8, &H60).Value)
    If cbrPopup Is Nothing Then
        Set ControlPopup = f.ComToolBar.Item(toolBar_PhotoStrong).FindControl(, ID_Active_SieveLens_Model, , True)
        ControlPopup.CommandBar.Controls.DeleteAll      '清空原有弹出菜单的内容
        
        '清空原来主菜单中的弹出菜单项
        Set MenuPopup = f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_SieveLens_Model, , True)
        MenuPopup.CommandBar.Controls.DeleteAll      '清空原有弹出菜单的内容
    Else
        Set ControlPopup = cbrPopup
    End If
    
    '增加新的弹出菜单
    For i = 1 To UBound(aPresetFilter)
        If UCase(aPresetFilter(i - 1).strModality) = strModality Then
            Set cbrToolBar = ControlPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_SieveLens_Model + i, aPresetFilter(i - 1).strname)
            cbrToolBar.Category = i - 1
            
            If Not MenuPopup Is Nothing Then
                Set cbrMenuBar = MenuPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_SieveLens_Model + i, aPresetFilter(i - 1).strname)
                cbrMenuBar.Category = i - 1
            End If
        End If
    Next i
End Sub

Public Sub subSetWidthLevelF(im As DicomImage, f As Form, Optional cbrPopup As CommandBarPopup)
'------------------------------------------------
'功能：设置窗宽窗位功能键弹出菜单
'参数： im--设置窗宽窗位的基准图像；
'       f--设置弹出菜单的窗体；
'       cbrPopup -- 为空则设置工具栏和主菜单的窗宽窗位菜单项；不为空，就是鼠标右键弹出菜单，则设置cbrPopup中的窗宽窗位菜单项；
'返回：无
'2009用
'-----------------------------------------------
    Dim strDriverType As String, intDriverType As Integer
    Dim i As Integer, j As Integer
    Dim ControlPopup As CommandBarPopup
    Dim cbrToolBar As CommandBarControl
    Dim cbrToolBarF2 As CommandBarControl
    Dim MenuPopup As CommandBarPopup    '主菜单中的弹出菜单项
    Dim cbrMenuBar As CommandBarControl '主菜单中的菜单项
    Dim cbrMenuBarF2 As CommandBarControl   '主菜单中的菜单项
    Dim blnIsMainViewer As Boolean          '是否主窗体，不是主窗体就是胶片打印窗体
    
    On Error GoTo err
    
    If im Is Nothing Then Exit Sub
    If IsNull(im.Attributes(&H8, &H60).Value) Then Exit Sub         '获取Modality
    If f.Name = "frmFilm" Then    '胶片打印窗体
        blnIsMainViewer = False
    Else
        blnIsMainViewer = True
    End If
    
    strDriverType = im.Attributes(&H8, &H60).Value
    If cbrPopup Is Nothing Then
        If blnIsMainViewer = False Then   '胶片打印窗体
            '清空工具栏中的弹出内容
            Set ControlPopup = f.CommBar_Film.Item(3).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow, , True)
            ControlPopup.CommandBar.Controls.DeleteAll
        Else
            '清空工具栏中的弹出菜单内容
            Set ControlPopup = f.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow, , True)
            ControlPopup.CommandBar.Controls.DeleteAll
            
            '清空主菜单中的弹出菜单内容
            Set MenuPopup = f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow, , True)
            MenuPopup.CommandBar.Controls.DeleteAll
        End If
    Else
        Set ControlPopup = cbrPopup
    End If
    
    intDriverType = 0
    
    For i = 1 To UBound(aPresetWinWL, 2)        '[找到图像对应设备]
        If UCase(aPresetWinWL(3, i).strModality) = UCase(strDriverType) Then
            intDriverType = i
            Exit For
        End If
    Next
    '''''''''''''''''''''''''''''''[增加F2菜单]''''''''''''''''''''''''''''''''''''''''
    Set cbrToolBarF2 = ControlPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_AdjustWindow_HandAdjustWindow_ReSet, "F2 自动")
    cbrToolBarF2.Checked = True
    If blnIsMainViewer Then
        If Not MenuPopup Is Nothing Then
            Set cbrMenuBarF2 = MenuPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_AdjustWindow_HandAdjustWindow_ReSet, "F2 自动")
            cbrMenuBarF2.Checked = True
        End If
        f.ComToolBar.KeyBindings.Add 0, VK_F2, ID_Active_AdjustWindow_HandAdjustWindow_ReSet
    Else
        f.CommBar_Film.KeyBindings.Add 0, VK_F2, ID_Active_AdjustWindow_HandAdjustWindow_ReSet
    End If
    
    ''''''''''''''''''''''''''''''[增加自定义按钮]'''''''''''''''''''''''''''''''''''''''''
    ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_Active_AdjustWindow_HandAdjustWindow_Custom, "自定义"
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
    If intDriverType > 0 Then
        For j = 3 To 12
            If aPresetWinWL(j, intDriverType).bInUse Then
                '增加窗宽窗位按钮
                Set cbrToolBar = ControlPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_AdjustWindow_HandAdjustWindow_F3 + j - 3, aPresetWinWL(j, intDriverType).strWinWLCName)
                cbrToolBar.Category = aPresetWinWL(j, intDriverType).lngWinWidth & "-" & aPresetWinWL(j, intDriverType).lngWinLevel
                If blnIsMainViewer Then
                    f.ComToolBar.KeyBindings.Add 0, VK_F3 + (j - 3), ID_Active_AdjustWindow_HandAdjustWindow_F3 + j - 3
                Else
                    f.CommBar_Film.KeyBindings.Add 0, VK_F3 + (j - 3), ID_Active_AdjustWindow_HandAdjustWindow_F3 + j - 3
                End If
                
                '增加主菜单的弹出菜单项
                If Not MenuPopup Is Nothing Then
                    Set cbrMenuBar = MenuPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_AdjustWindow_HandAdjustWindow_F3 + j - 3, aPresetWinWL(j, intDriverType).strWinWLCName)
                    cbrMenuBar.Category = aPresetWinWL(j, intDriverType).lngWinWidth & "-" & aPresetWinWL(j, intDriverType).lngWinLevel
                End If
                '设置默认按钮
                If aPresetWinWL(j, intDriverType).intDefault = 1 Then
                    cbrToolBarF2.Checked = False
                    cbrToolBar.Checked = True
                    
                    If Not MenuPopup Is Nothing Then
                        cbrMenuBarF2.Checked = False
                        cbrMenuBar.Checked = True
                    End If
                End If
            End If
        Next
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetAngle(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal X3 As Double, ByVal Y3 As Double) As Double
'------------------------------------------------
'功能：计算通过三个点连成的两条线之间的角度
'参数：（X1,Y1）－－两直线交点的X,Y坐标；（X2,Y2）－－直线1上的一点；（X3,Y3）－－直线2上的一点
'返回：GetAngle－－两条直线之间的角度，单位为：度
'2009用
'------------------------------------------------
    Dim Pi As Double
    Dim dblCos As Double, dblAngle1 As Double, dblAngle2 As Double
    Pi = 3.14159265358979
    If x1 = x2 And y1 = y2 Then
        dblAngle1 = 0
    Else
        dblCos = (x2 - x1) / Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
        If Abs(dblCos) = 1 Then
            dblAngle1 = IIf(dblCos = 1, 0, Pi)
        Else
            dblAngle1 = Atn(-dblCos / Sqr(-dblCos * dblCos + 1)) + 2 * Atn(1)
        End If
    End If
    If x1 = X3 And y1 = Y3 Then
        dblAngle2 = 0
    Else
        dblCos = (X3 - x1) / Sqr((X3 - x1) ^ 2 + (Y3 - y1) ^ 2)
        If Abs(dblCos) = 1 Then
            dblAngle2 = IIf(dblCos = 1, 0, Pi)
        Else
            dblAngle2 = Atn(-dblCos / Sqr(-dblCos * dblCos + 1)) + 2 * Atn(1)
        End If
    End If
    GetAngle = IIf((y2 - y1) * (Y3 - y1) > 0, Abs(dblAngle1 - dblAngle2), Abs(dblAngle1 + dblAngle2)) * 180 / Pi
    If GetAngle > 180 Then GetAngle = 360 - GetAngle
End Function

Function Max7InArray(a As Variant, ByRef lMax As Long, ByRef lMin As Long) As Long
'------------------------------------------------
'功能：从数组里面取值最大的7个下标，对其求平均值
'参数：a--进行操作的数组；lMax--用来返回最大下标   lMin--用来返回最小下标。
'返回：返回值为此平均值，并通过对lMax和lMin来返回最大和最小下标。
'2009用
'------------------------------------------------
    Dim m1 As Long, m2 As Long, m3 As Long, m4 As Long, m5 As Long, m6 As Long, m7 As Long
    Dim c1 As Long, c2 As Long, c3 As Long, c4 As Long, c5 As Long, c6 As Long, c7 As Long
    Dim s As Long
    Dim cMax As Long
    Dim cMin As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cMax = 1
    cMin = 1
    Max7InArray = 0
    m1 = 0
    m2 = 0
    m3 = 0
    m4 = 0
    m5 = 0
    m6 = 0
    m7 = 0
    s = a(1)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim lCount As Long
    lCount = 1
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    While lCount <= UBound(a)
        If a(lCount) > m7 Then
           If a(lCount) > m6 Then
               If a(lCount) > m5 Then
                    If a(lCount) > m4 Then
                       If a(lCount) > m3 Then
                           If a(lCount) > m2 Then
                               If a(lCount) > m1 Then
                                   m1 = a(lCount)
                                   cMax = lCount
                                   c1 = lCount
                               Else
                                   m2 = a(lCount)
                                   c2 = lCount
                               End If
                           Else
                               m3 = a(lCount)
                               c3 = lCount
                           End If
                       Else
                           m4 = a(lCount)
                           c4 = lCount
                       End If
                    Else
                       m5 = a(lCount)
                       c5 = lCount
                    End If
               Else
                    m6 = a(lCount)
                    c6 = lCount
               End If
           Else
                m7 = a(lCount)
                c7 = lCount
           End If
        End If
        ''''''判断小值'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If a(lCount) < s Then
            s = a(lCount)
            cMin = lCount
        End If
        lCount = lCount + 1
    Wend
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Max7InArray = (c1 + c2 + c3 + c4 + c5 + c6 + c7) / 7
    lMax = cMax
    lMin = cMin
End Function

Private Function funIsShutter(strModality As String) As Integer
'------------------------------------------------
'功能：判断输入的影像类别是否需要进行图像消隐操作
'参数：
'     strModality 进行图像消隐的影像类别
'返回：0-没有设置，不用处理；1－无消隐；2－有图像消隐。
'2009用
'------------------------------------------------
    Dim i As Integer
    funIsShutter = 0
    For i = 1 To UBound(aImageShutter)
        If UCase(aImageShutter(i).strModality) = UCase(strModality) Then
            If aImageShutter(i).intShutterType > 0 And aImageShutter(i).intShutterType < 8 Then
                funIsShutter = 2    '有图像消隐
            Else
                funIsShutter = 1    '无图像消隐
            End If
        End If
    Next i
End Function

Public Sub subDrawImgShutter(img As DicomImage, Optional isForce As Boolean = False)
'------------------------------------------------
'功能：根据系统设置的影像类别，给输入的图像画图像消隐
'参数：
'       img 进行图像消隐的图像
'       isForce －当被设置为不使用消隐的时候，是否强制删图像中除现有的消隐信息
'返回：无
'2009用
'------------------------------------------------
    Dim iResult As Integer
    Dim intModality As Integer
    Dim strModality As String
    Dim i As Integer
    Dim strArray() As String
    Dim intCount As Integer
    Dim intShutterType As Integer
    Dim strCenter(2) As String
    Dim strVertices() As String
    
    If IsNull(img.Attributes(&H8, &H60).Value) Then Exit Sub
    strModality = img.Attributes(&H8, &H60).Value
    iResult = funIsShutter(strModality)
    If isForce = False And iResult = 0 Then Exit Sub
    For i = 1 To UBound(aImageShutter)
        If UCase(aImageShutter(i).strModality) = UCase(strModality) Then
            intModality = i
            Exit For
        End If
    Next i
    
    '处理图像消隐
    If aImageShutter(intModality).intShutterType > 0 And aImageShutter(intModality).intShutterType < 8 Then
        '处理消隐类型
        intShutterType = aImageShutter(intModality).intShutterType
        intCount = 0
        If intShutterType >= 4 Then     '多边形消隐
            intShutterType = intShutterType - 4
            intCount = intCount + 1
            ReDim Preserve strArray(intCount) As String
            strArray(intCount) = "POLYGONAL"
            '大于三个点，且是偶数，则将多边形顶点添加到图像中
            strVertices = Split(aImageShutter(intModality).strVertices, ":")
            If UBound(strVertices) >= 5 And UBound(strVertices) Mod 2 = 1 Then
                ReDim Preserve strVertices(UBound(strVertices) + 1) As String
                For i = UBound(strVertices) To 1 Step -1
                    strVertices(i) = strVertices(i - 1)
                Next i
                img.Attributes.Add &H18, &H1620, strVertices
            End If
        End If
        If intShutterType >= 2 Then     '矩形消隐
            intShutterType = intShutterType - 2
            intCount = intCount + 1
            ReDim Preserve strArray(intCount) As String
            strArray(intCount) = "RECTANGULAR"
            img.Attributes.Add &H18, &H1602, aImageShutter(intModality).intRectLeft
            img.Attributes.Add &H18, &H1604, aImageShutter(intModality).intRectRight
            img.Attributes.Add &H18, &H1606, aImageShutter(intModality).intRectUpper
            img.Attributes.Add &H18, &H1608, aImageShutter(intModality).intRectLower
        End If
        If intShutterType >= 1 Then     '圆形消隐
            intCount = intCount + 1
            ReDim Preserve strArray(intCount) As String
            strArray(intCount) = "CIRCULAR"
            '添加圆心和半径
            strCenter(1) = aImageShutter(intModality).intCenterX
            strCenter(2) = aImageShutter(intModality).intCenterY
            img.Attributes.Add &H18, &H1610, strCenter
            img.Attributes.Add &H18, &H1612, aImageShutter(intModality).intRadius
        End If
        img.Attributes.Add &H18, &H1600, strArray
        img.Attributes.Add &H18, &H1622, aImageShutter(intModality).lngColor
    Else        '处理“无消隐“
        img.Attributes.Remove &H18, &H1600
        img.Attributes.Remove &H18, &H1602
        img.Attributes.Remove &H18, &H1604
        img.Attributes.Remove &H18, &H1606
        img.Attributes.Remove &H18, &H1608
        img.Attributes.Remove &H18, &H1610
        img.Attributes.Remove &H18, &H1612
        img.Attributes.Remove &H18, &H1620
        img.Attributes.Remove &H18, &H1622
    End If
    img.Refresh False
End Sub

Public Function funGetLinePoints(img As DicomImage, la As DicomLabel, aGrey() As Integer, intBeginX As Integer _
                , intBeginY As Integer, intEndX As Integer, intEndY As Integer) As Boolean
'------------------------------------------------
'功能：从图像的给定线型标注（直线、折线）中提取灰度值数组和起点、终点坐标
'参数：
'       img －包含标注的图像。
'       la － 提取灰度值的线型标注。
'       aGrey－保存灰度值的数组，返回值。
'       intBeginX－起点的X坐标，返回值。
'       intBeginY －起点的Y坐标，返回值。
'       intEndX－终点的X坐标，返回值。
'       intEndY－终点的X坐标，返回值。
'返回：是否成功返回了灰度值数组。True－正常返回。Fasle－执行失败，可能是标注不是线型标注。
'2009用
'------------------------------------------------
    '获取直线上灰度值，存放到数组中
    Dim vPixels As Variant
    Dim i As Integer
    Dim iFrame As Integer
    Dim lngCount As Long        '点的数量
    Dim iSizex As Integer       '图像的x方向点数
    Dim iSizey As Integer       '图像的y方向点数
    Dim iTempx As Integer       '当前的临时x点
    Dim iTempy As Integer       '当前的临时y点
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    iFrame = img.Frame          '记录当前图像的侦数
    vPixels = img.Pixels        '获取当前图像的像素点
    iSizex = img.sizex          '获取当前图像的x方向点数
    iSizey = img.sizey          '获取当前图像的y方向点数
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If la.width = 0 And la.height = 0 Then  '宽度、高度都为零，则设置宽度为1
        intBeginX = la.left
        intEndX = la.left + 1
        intBeginY = la.top
        intEndY = la.top
        lngCount = intEndX - intBeginX + 1
        ReDim aGrey(lngCount) As Integer
        aGrey(lngCount - 1) = vPixels(intBeginX, intBeginY, iFrame)
        aGrey(lngCount) = vPixels(intEndX, intBeginY, iFrame)
        funGetLinePoints = True
        Exit Function
    End If
    '分成两种情况填充灰度数组
    If la.LabelType = doLabelLine Then      ' 对于直线的操作
        Dim lngW As Long
        Dim lngH As Long
        Dim iCount As Long
        iCount = 1
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        lngW = Abs(la.width) + 1
        lngH = Abs(la.height) + 1
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If lngW > lngH Then             '宽度大于高度，按照x方向递增，从左到右计算直线上的点
            If la.width < 0 Then        'left在右边，需要调换顺序
                intEndX = la.left
                intEndY = la.top
                intBeginX = la.left + la.width
                intBeginY = la.top + la.height
            Else                        'left在左边，begin点直接取left,top点
                intBeginX = la.left
                intBeginY = la.top
                intEndX = la.left + la.width
                intEndY = la.top + la.height
            End If
    
            '确保intBeginX到intEndX的值在1到图像的sizex之间
'            If intBeginX < 1 Then intBeginX = 1
'            If intBeginX > iSizex Then intBeginX = iSizex
'            If intEndX < 1 Then intEndX = 1
'            If intEndX > iSizex Then intEndX = iSizex
    
            lngCount = intEndX - intBeginX + 1
            ReDim aGrey(lngCount) As Integer
    
            For i = intBeginX To intEndX
                iTempx = i
                iTempy = la.height / la.width * (i - intBeginX) + intBeginY
                '确保iTempx的值在1到图像的sizex之间
                If iTempx < 1 Then iTempx = 1
                If iTempx > iSizex Then iTempx = iSizex
                '确保iTempy的值在1到图像的sizey之间
                If iTempy < 1 Then iTempy = 1
                If iTempy > iSizey Then iTempy = iSizey
                aGrey(iCount) = vPixels(iTempx, iTempy, iFrame)
                iCount = iCount + 1
            Next
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Else                            '高度大于宽度，按照y方向递增，从上到下计算直线上的点
            If la.height < 0 Then       'top在下边，需要调换顺序
                intEndX = la.left
                intEndY = la.top
                intBeginX = la.left + la.width
                intBeginY = la.top + la.height
            Else                        'top在上边，begin点直接取left,top点
                intBeginX = la.left
                intBeginY = la.top
                intEndX = la.left + la.width
                intEndY = la.top + la.height
            End If
    
            '确保intBeginY到intEndY的值在1到图像的sizey之间
            If intBeginY < 1 Then intBeginY = 1
            If intBeginY > iSizey Then intBeginY = iSizey
            If intEndY < 1 Then intEndY = 1
            If intEndY > iSizey Then intEndY = iSizey
    
            lngCount = intEndY - intBeginY + 1
            ReDim aGrey(lngCount) As Integer
            For i = intBeginY To intEndY
                iTempx = la.width / la.height * (i - intBeginY) + intBeginX
                iTempy = i
                '确保iTempx的值在1到图像的sizex之间
                If iTempx < 1 Then iTempx = 1
                If iTempx > iSizex Then iTempx = iSizex
                '确保iTempy的值在1到图像的sizey之间
                If iTempy < 1 Then iTempy = 1
                If iTempy > iSizey Then iTempy = iSizey
                aGrey(iCount) = vPixels(iTempx, iTempy, iFrame)
                iCount = iCount + 1
            Next
        End If
        funGetLinePoints = True
    Else                            '对于多边线，直接对Points操作
        Dim vPoints As Variant

        vPoints = la.Points
        lngCount = UBound(vPoints) / 2
        ReDim aGrey(lngCount) As Integer
        For i = 1 To lngCount
            iTempx = vPoints(2 * i - 1)
            iTempy = vPoints(2 * i)
            
            '确保iTempx的值在1到图像的sizex之间
            If iTempx < 1 Then iTempx = 1
            If iTempx > iSizex Then iTempx = iSizex
            
            '确保iTempy的值在1到图像的sizey之间
            If iTempy < 1 Then iTempy = 1
            If iTempy > iSizey Then iTempy = iSizey
            
            aGrey(i) = vPixels(iTempx, iTempy, iFrame)
        Next
        intBeginX = vPoints(1)
        intBeginY = vPoints(2)
        intEndX = vPoints(lngCount * 2 - 1)
        intEndY = vPoints(lngCount * 2)
        funGetLinePoints = True
    End If
End Function

Public Function funGetVasEdge(img As DicomImage, lblLine As DicomLabel, intThreshold As Integer, x1 As Long, y1 As Long, x2 As Long, y2 As Long) As Boolean
'------------------------------------------------
'功能：用于血管狭窄测量，根据直线标注和预设的阈值，查找血管壁的坐标。
'参数：
'       img －包含标注的图像。
'       lblLine － 垂直于血管的直线标注。
'       (x1,y1)－左边血管壁跟血管垂直线交点的坐标，返回值。
'       (x2,y2)－右边血管壁跟血管垂直线交点的坐标，返回值。
'返回：是否成功计算出了两个血管壁坐标。True－正常返回。Fasle－执行失败，可能输入的直线标注不是线型标注。
'2009用
'------------------------------------------------
    Dim aGrey() As Integer
    Dim intBeginX As Integer, intBeginY As Integer, intEndX As Integer, intEndY As Integer
    Dim lngCount As Long
    Dim i As Integer
    Dim lngCenter As Long
    Dim intLower As Integer
    Dim intUpper As Integer
    
    If lblLine.LabelType <> doLabelLine Then Exit Function
    'If Abs(lblLine.width) < 2 And Abs(lblLine.height) < 2 Then Exit Function
    If funGetLinePoints(img, lblLine, aGrey, intBeginX, intBeginY, intEndX, intEndY) = False Then Exit Function
    lngCount = UBound(aGrey)
    lngCenter = lngCount \ 2
    '往图像左上角找血管壁
    intLower = 1        '初始化左上角血管壁
    For i = lngCenter To 1 Step -1
        If Abs(aGrey(i) - aGrey(lngCenter)) > intThreshold Then
            intLower = i
            Exit For
        End If
    Next i
    '往图像右下角找血管壁
    intUpper = lngCount     '初始化右下角血管壁
    For i = lngCenter + 1 To lngCount Step 1
        If Abs(aGrey(i) - aGrey(lngCenter)) > intThreshold Then
            intUpper = i
            Exit For
        End If
    Next i
    '判断直线的斜率，是按照X来计算，还是按照Y来计算
    If lngCount = intEndY - intBeginY + 1 Then '按照Y来计算
        y1 = intBeginY + intLower
        y2 = intBeginY + intUpper
        x1 = (y1 - intEndY) / (intBeginY - intEndY) * (intBeginX - intEndX) + intEndX
        x2 = (y2 - intEndY) / (intBeginY - intEndY) * (intBeginX - intEndX) + intEndX
    Else        '按照X来计算
        x1 = intBeginX + intLower
        x2 = intBeginX + intUpper
        y1 = (x1 - intEndX) / (intBeginX - intEndX) * (intBeginY - intEndY) + intEndY
        y2 = (x2 - intEndX) / (intBeginX - intEndX) * (intBeginY - intEndY) + intEndY
    End If
    funGetVasEdge = True
End Function

Public Sub subDrawVasEdgeLine(lblLine As DicomLabel, lblShortLine As DicomLabel, intCenterX As Long, intCenterY As Long)
'------------------------------------------------
'功能：用于血管狭窄测量，根据直线标注和血管壁的坐标，确定并画出血管壁短直线的位置。
'参数：
'       lblLine － 垂直于血管的直线标注。
'       lblShortLine － 血管壁短直线。
'       (intCenterX,intCenterY)－血管壁跟血管垂直线交点的坐标。
'返回：无，直接移动血管壁短直线的位置。
'2009用
'------------------------------------------------
    Dim lngLineWidth As Long
    Dim intNewX As Integer
    Dim intNewY As Integer
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    lngWidth = Abs(lblLine.width)
    lngHeight = Abs(lblLine.height)
    lngLineWidth = Sqr(lngHeight * lngHeight + lngWidth * lngWidth)
    If lngLineWidth = 0 Then
        intNewX = 0
        intNewY = 0
    Else
        intNewX = lngHeight / lngLineWidth * intVasEdgeWidth / 2
        intNewY = lngWidth / lngLineWidth * intVasEdgeWidth / 2
    End If
    
    If (lblLine.width > 0 And lblLine.height > 0) Or (lblLine.width < 0 And lblLine.height < 0) Then
        intNewY = -intNewY
    End If
    lblShortLine.left = intCenterX - intNewX
    lblShortLine.top = intCenterY - intNewY
    lblShortLine.height = intNewY * 2
    lblShortLine.width = intNewX * 2
End Sub

Public Sub subCenterZoom(img As DicomImage, Viewer As DicomViewer, dblZoom As Double)
'------------------------------------------------
'功能：对图像进行缩放。以当前viewer中心点为缩放中心点。
'参数：
'       img -- 进行缩放的图像
'       viewer －－ 图像所在的viewer
'       dblZoom －－图像新的缩放倍数
'返回：无，直接调整图像的缩放倍数
'2009用
'------------------------------------------------
    Dim dblOldZoom As Double
    Dim lngOldScroX As Long
    Dim lngOldScroY As Long
    Dim dblZoomRatio As Double
    
    On Error GoTo err
    
    If img Is Nothing Then Exit Sub
    If img.ActualZoom = 0 Then Exit Sub
    
    dblOldZoom = img.ActualZoom
    lngOldScroX = img.ActualScrollX
    lngOldScroY = img.ActualScrollY
    img.Zoom = dblZoom
    img.StretchToFit = False
    
    dblZoomRatio = 1 - img.ActualZoom / dblOldZoom
    img.ScrollX = lngOldScroX - (lngOldScroX + Viewer.width / Viewer.MultiColumns / Screen.TwipsPerPixelX / 2) * dblZoomRatio
    img.ScrollY = lngOldScroY - (lngOldScroY + Viewer.height / Viewer.MultiRows / Screen.TwipsPerPixelY / 2) * dblZoomRatio
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub RectangleZoom(FilmViewer As DicomViewer, img As DicomImage, lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long)
'------------------------------------------------
'功能：对图像进行缩放。按照Viewer框的大小缩放
'参数：
'       img -- 进行缩放的图像
'       viewer －－ 图像所在的viewer
'       dblZoom －－图像新的缩放倍数
'返回：无，直接调整图像的缩放倍数
'2009用
'------------------------------------------------
    Dim newZoom As Double
    Dim dblRatio As Double
    Dim oldZoom As Double
    
    If lngWidth > 0 And lngHeight > 0 Then
        oldZoom = img.ActualZoom
        img.StretchToFit = False

        dblRatio = FilmViewer.width / FilmViewer.MultiColumns / Screen.TwipsPerPixelX / lngWidth
        If dblRatio > FilmViewer.height / FilmViewer.MultiRows / Screen.TwipsPerPixelY / lngHeight Then
            dblRatio = FilmViewer.height / FilmViewer.MultiRows / Screen.TwipsPerPixelY / lngHeight
        End If
        
        newZoom = oldZoom * dblRatio
        img.Zoom = newZoom
        
        img.ScrollX = lngLeft * dblRatio
        img.ScrollY = lngTop * dblRatio
    End If
End Sub

Public Function CutOutAImage(img As DicomImage)
    Dim Simg As New DicomImage
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lblFrame As DicomLabel
    Dim lngTemp As Long
    
    Set CutOutAImage = Simg
    On Error GoTo err
    
    Set lblFrame = img.Labels(img.Labels.Count)
    
    '图象框的位置
    If lblFrame.width >= 0 Then
        lngLeft = lblFrame.left
    Else
        lngLeft = lblFrame.left + lblFrame.width
    End If
    lngWidth = Abs(lblFrame.width)
    
    If lblFrame.height >= 0 Then
        lngTop = lblFrame.top
    Else
        lngTop = lblFrame.top + lblFrame.height
    End If
    lngHeight = Abs(lblFrame.height)
    
    lblFrame.Visible = False
    
    '裁剪图像
    If (img.RotateState = doRotateLeft And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipBoth) Then
        '左旋,右旋+全镜
        lngLeft = img.sizex - lngLeft - lngWidth
        
        lngTemp = lngLeft
        lngLeft = lngTop
        lngTop = lngTemp
    ElseIf (img.RotateState = doRotateRight And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateLeft And img.FlipState = doFlipBoth) Then
        '右旋 ,左旋+全镜
        lngTop = img.sizey - lngTop - lngHeight
        
        lngTemp = lngLeft
        lngLeft = lngTop
        lngTop = lngTemp
    ElseIf (img.RotateState = doRotate180 And img.FlipState = doFlipNormal) _
         Or (img.RotateState = doRotateNormal And img.FlipState = doFlipBoth) Then
         '180度 和 全镜像
        lngLeft = img.sizex - lngLeft - lngWidth
        lngTop = img.sizey - lngTop - lngHeight
    ElseIf (img.RotateState = doRotateNormal And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipVertical) Then
        '左右镜像,180+上下倒置
        lngLeft = img.sizex - lngLeft - lngWidth
    ElseIf (img.RotateState = doRotateNormal And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipHorizontal) Then
        '上下倒置,180+左右
        lngTop = img.sizey - lngTop - lngHeight
    ElseIf (img.RotateState = doRotateRight And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateLeft And img.FlipState = doFlipVertical) Then
        '左旋+左右镜像！这里左右是反的
        '右旋+上下倒置！这里左右是反的
        lngTop = img.sizey - lngTop - lngHeight
        lngLeft = img.sizex - lngLeft - lngWidth
        
        lngTemp = lngLeft
        lngLeft = lngTop
        lngTop = lngTemp
    ElseIf (img.RotateState = doRotateLeft And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipVertical) Then
        '右旋+左右镜像！ 这里左右是反的
        '左旋+上下倒置！ 这里左右是反的
        lngTemp = lngLeft
        lngLeft = lngTop
        lngTop = lngTemp
    End If
    Set Simg = img.SubImage(lngLeft, lngTop, lngWidth, lngHeight, 1, 1)
    
    Set CutOutAImage = Simg
    Exit Function
err:

End Function


Public Sub subWriteDicomPara(imgSource As DicomImage, imgDest As DicomImage)
'------------------------------------------------
'功能：给输入的图像填写DICOM文件头信息
'参数：img－－输入的DICOM文件,lngAdviceID－－医嘱ID
'返回：无，直接文件头信息写入img的文件头
'------------------------------------------------
    Dim curDate As Date
    Dim attr As DicomAttribute
    Dim Dicomglb As New DicomGlobal

    curDate = zlDatabase.Currentdate
    
    imgDest.InstanceUID = Dicomglb.NewUID
    imgDest.StudyUID = imgSource.StudyUID
    imgDest.SeriesUID = imgSource.SeriesUID
    
    imgDest.Attributes.Add &H8, &H8, ""                             'ImageType  空
    imgDest.Attributes.Add &H8, &H16, "1.2.840.10008.5.1.4.1.1.7"   'SOP Class  UID，二次捕捉
    imgDest.Attributes.Add &H8, &H20, Format(curDate, "yyyy-mm-dd")     'Study Date 检查日期
    imgDest.Attributes.Add &H8, &H21, Format(curDate, "yyyy-mm-dd")     'Series Date 序列日期
    imgDest.Attributes.Add &H8, &H22, Format(curDate, "yyyy-mm-dd")     'Acquisition Date 采集日期
    imgDest.Attributes.Add &H8, &H23, Format(curDate, "yyyy-mm-dd")     'Image Date   图像日期
    imgDest.Attributes.Add &H8, &H30, Format(curDate, "HH24:MI:SS")     'Study Time   检查时间
    imgDest.Attributes.Add &H8, &H31, Format(curDate, "HH24:MI:SS")     'Series Time  序列时间
    imgDest.Attributes.Add &H8, &H32, Format(curDate, "HH24:MI:SS")     'Acquisition Time  采集时间
    imgDest.Attributes.Add &H8, &H33, Format(curDate, "HH24:MI:SS")     'Image Time  图像时间
    imgDest.Attributes.Add &H8, &H50, ""                            'Accession Number 空
    imgDest.Attributes.Add &H8, &H60, imgSource.Attributes(&H8, &H60).Value                  'Modality 影像类别
    imgDest.Attributes.Add &H8, &H70, "ZLSOFT"                      'Manufacturer 厂商
    imgDest.Attributes.Add &H8, &H80, "ZLSOFT"                'Institution Name 单位名称
    imgDest.Attributes.Add &H8, &H90, ""                            'Referring Physician's Name 空
'    imgDest.Attributes.Add &H8, &H1030, ""                          'Study Description 检查描述 空
    imgDest.Attributes.Add &H10, &H10, imgSource.Name                       'Name 姓名
    imgDest.Attributes.Add &H10, &H20, imgSource.PatientID                 'Patient ID 病人ID
    imgDest.Attributes.Add &H10, &H30, imgSource.DateOfBirth                  'BirthDate 生日
    imgDest.Attributes.Add &H10, &H40, imgSource.Sex                        'Sex 性别
    imgDest.Attributes.Add &H10, &H4000, ""                         'Patient Comment 病人注释
    imgDest.Attributes.Add &H20, &H10, "1"                   'Study ID 检查ID
    imgDest.Attributes.Add &H20, &H11, "1"                          'Series Number 序列号
    imgDest.Attributes.Add &H20, &H13, "1"                          'ImageNumber 图像号
    imgDest.Attributes.Add &H20, &H20, ""                           'Orientation 空
    
'    添加标尺信息
    If imgSource.Attributes(&H28, &H30).Exists Then
        imgDest.Attributes.Add &H28, &H30, imgSource.Attributes(&H28, &H30).Value
    End If
    'KODAK CR800 使用以下的标尺信息
    If imgSource.Attributes(&H18, &H1164).Exists Then
        imgDest.Attributes.Add &H18, &H1164, imgSource.Attributes(&H18, &H1164).Value
    End If
End Sub

Public Function funCopyMPRControlLines(im As DicomImage, oldImage As DicomImage)
'------------------------------------------------
'功能：初始化指定图像的矢冠状重建控制点和控制线
'参数：     im--进行矢冠状重建的轴位图像
'           oldImage -- 需要拷贝控制线的原图，可能是空
'返回：无，直接对im图像上的矢冠状重建标注做初始化。
'------------------------------------------------
    On Error GoTo err
    
    If oldImage Is Nothing Then
        Exit Function
    End If
    
    im.Labels(G_INT_SYS_LABEL_MPRV).left = oldImage.Labels(G_INT_SYS_LABEL_MPRV).left
    im.Labels(G_INT_SYS_LABEL_MPRV).top = oldImage.Labels(G_INT_SYS_LABEL_MPRV).top
    im.Labels(G_INT_SYS_LABEL_MPRV).height = oldImage.Labels(G_INT_SYS_LABEL_MPRV).height
    im.Labels(G_INT_SYS_LABEL_MPRV).width = oldImage.Labels(G_INT_SYS_LABEL_MPRV).width
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPRH).left = oldImage.Labels(G_INT_SYS_LABEL_MPRH).left
    im.Labels(G_INT_SYS_LABEL_MPRH).top = oldImage.Labels(G_INT_SYS_LABEL_MPRH).top
    im.Labels(G_INT_SYS_LABEL_MPRH).height = oldImage.Labels(G_INT_SYS_LABEL_MPRH).height
    im.Labels(G_INT_SYS_LABEL_MPRH).width = oldImage.Labels(G_INT_SYS_LABEL_MPRH).width
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_V1).left = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V1).left
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_V1).top = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V1).top
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_V2).left = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V2).left
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_V2).top = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V2).top
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_H1).left = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H1).left
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_H1).top = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H1).top
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_H2).left = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H2).left
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_H2).top = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H2).top
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Refresh False
     
    funCopyMPRControlLines = True
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function funImageSmoothing(ByRef img() As Integer, intTimes As Integer) As Boolean
'------------------------------------------------
'功能：对二维数组中的图像，做平滑处理
'参数： img() -- 图像二维数组
'       intTimes -- 平滑次数，一般是1-2次
'返回：无，返回重建后的图像
'------------------------------------------------
    Dim i As Long, j As Long, k As Long
    Dim TolHeight As Long       '图像高度
    Dim lngWidth As Long
    Dim lngTemp As Long          '为了防止integer类型的a,b数组内容在平滑过程中的溢出，使用long型的中间变量
    
    On Error GoTo err
    
    TolHeight = UBound(img, 2)
    lngWidth = UBound(img, 1)
    
    For k = 1 To intTimes
        For i = 3 To TolHeight - 2
            For j = 2 To lngWidth - 1
                '这个模板很好，对细节处理很锐利
                lngTemp = CLng(img(j - 1, i - 1)) + CLng(img(j - 1, i + 1)) + _
                          CLng(2 * img(j, i - 2)) + CLng(2 * img(j, i - 1)) + CLng(img(j, i))
                lngTemp = (CLng(lngTemp) + CLng(2 * img(j, i + 1)) + CLng(2 * img(j, i + 2)) _
                         + CLng(img(j + 1, i - 1)) + CLng(img(j + 1, i + 1)) _
                         ) / 13
                img(j, i) = lngTemp
            Next
        Next
        '''''对最上两根线做处理''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For j = 2 To lngWidth - 1
            lngTemp = CLng(img(j - 1, 1)) + CLng(img(j - 1, 2)) + CLng(2 * img(j - 1, 3)) + CLng(img(j, 1)) + CLng(img(j, 2))
            lngTemp = (CLng(lngTemp) + CLng(2 * img(j, 3)) + CLng(img(j + 1, 1)) + CLng(img(j + 1, 2)) + CLng(2 * img(j + 1, 3))) / 12
            img(j, 2) = lngTemp
        Next
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For j = 1 To lngWidth - 1
            lngTemp = CLng(img(j - 1, 1)) + CLng(img(j - 1, 2)) + CLng(img(j - 1, 3)) + CLng(img(j, 1))
            lngTemp = (CLng(lngTemp) + CLng(img(j, 2)) + CLng(img(j, 3)) + CLng(img(j + 1, 1)) + CLng(img(j + 1, 2)) + CLng(img(j + 1, 3))) / 9
            img(j, 1) = lngTemp
        Next
        ''''''对最下两根线做处理''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For j = 2 To lngWidth - 1
            lngTemp = CLng(2 * img(j - 1, TolHeight - 2)) + CLng(img(j - 1, TolHeight - 1)) + CLng(img(j - 1, TolHeight)) + CLng(2 * img(j, TolHeight - 2))
            lngTemp = (CLng(lngTemp) + CLng(img(j, TolHeight - 1)) + CLng(img(j, TolHeight)) + CLng(2 * img(j + 1, TolHeight - 2)) + CLng(img(j + 1, TolHeight - 1)) + CLng(img(j + 1, TolHeight))) / 12
            img(j, TolHeight - 1) = lngTemp
        Next
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For j = 1 To lngWidth - 1
            lngTemp = CLng(img(j - 1, TolHeight)) + CLng(img(j - 1, TolHeight - 1)) + CLng(img(j - 1, TolHeight - 2)) + CLng(img(j, TolHeight)) + CLng(img(j, TolHeight - 1))
            lngTemp = (CLng(lngTemp) + CLng(img(j, TolHeight - 2)) + CLng(img(j + 1, TolHeight)) + CLng(img(j + 1, TolHeight - 1)) + CLng(img(j + 1, TolHeight - 2))) / 9
            img(j, TolHeight) = lngTemp
        Next
   Next
   
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

