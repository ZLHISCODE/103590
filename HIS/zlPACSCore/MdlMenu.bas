Attribute VB_Name = "MdlMenu"
Option Explicit
'--------------------------------------------------------
'功  能：本模块为菜单按钮调用的处理部分
'编制人：胡涛
'编制日期：2004.6.12
'过程函数清单：
'    subFakeColor():            对窗体中选中的图像做伪彩处理，显示一个窗体，让用户选择需要的伪彩方案。
'    subFunctionWL():           功能键设置窗宽窗位处理
'    subFilm():                 胶片打印
'    subcalibrate():            校准
'    SubImageUnsharp():         图像增强
'    subMnuImageSort():         排序方式处理，根据lngToolID来处理排序，同时处理菜单的选择状态
'    subMouseRLset():           处理鼠标左右键的check状态。
'    subCurrentCheck():         定位线处理，处理定位线菜单的单击事件，控制定位线三个相关按钮只能被选中一个
'    subOutputToPowerPoint():   输出到POWERPOINT
'    subDSA():                  DSA数字减影
'    subCutOut():               进入和退出裁减状态，隐藏或显示裁剪标注
'    subDispLabelInfo():        显示或隐藏图像的用户标注信息
'    subManipulation():         多平面处理（图像旋转、反白等）
'    subSelectAllSerial():      选择所有序列
'    subSelectAllIMage():       选择所有图像
'    subFullScreen():           切换屏幕的全屏状态
'修改记录：
'    2005.07.08    黄捷
'-------------------------------------------------------


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''
''''
''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub subFakeColor(f As frmViewer)
'------------------------------------------------
'功能：对窗体中选中的图像做伪彩处理，显示一个窗体，让用户选择需要的伪彩方案。
'参数：f--进行伪彩处理的窗体。
'返回：无
'2009用
'------------------------------------------------
    If f.intSelectedSerial < 1 Then Exit Sub
    Dim strSQL As String, rsTemp As Recordset
    Set FrmFakeColor.f = f
    strSQL = "SELECT 颜色,序号,系统方案 FROM  影像颜色清单"
    If blLocalRun = True Then
        Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询伪彩颜色")
    End If
    Do While Not rsTemp.EOF
        FrmFakeColor.cobColor.AddItem IIf(rsTemp!系统方案 = 1, "系统方案：", "用户方案：") & rsTemp!颜色
        FrmFakeColor.cobColor.ItemData(FrmFakeColor.cobColor.NewIndex) = rsTemp!序号
        rsTemp.MoveNext
    Loop
    rsTemp.MoveFirst
    FrmFakeColor.cobColor.ListIndex = 0
    FrmFakeColor.Show 1, f
End Sub

Public Sub subFunctionFilter(ByVal control As CommandBarControl, f As frmViewer)
'------------------------------------------------
'功能：功能键设置滤镜模板处理
'参数： Control--菜单控件；
'       f--窗体。
'返回：无
'------------------------------------------------
    Dim iRow As Integer
    Dim dblTemp As Double
    
    On Error GoTo err
    If f.SelectedImage Is Nothing Then Exit Sub
    If control Is Nothing Then Exit Sub
    If control.Id >= ID_Active_SieveLens_Model + 1 And control.Id < ID_Active_SieveLens_Model + 40 Then
        iRow = Val(control.Category)
        If iRow >= 0 And iRow < UBound(aPresetFilter) Then
            f.SelectedImage.UnsharpEnhancement = 0
            f.SelectedImage.UnsharpLength = 0
            f.SelectedImage.FilterLength = 0
            
            '图像处理，计算增强值
            '图像增强强度增加
            If aPresetFilter(iRow).intUnSharpEnhancementUp > 0 Then
                Call SubImageFiltering("miUnSharpEnhancementUp", f.SelectedImage, aPresetFilter(iRow).intUnSharpEnhancementUp)
            End If
            '图像增强强度减少
            If aPresetFilter(iRow).intUnSharpEnhancementDown > 0 Then
                Call SubImageFiltering("miUnSharpEnhancementDown", f.SelectedImage, aPresetFilter(iRow).intUnSharpEnhancementDown)
            End If
            
            '图像增强幅度增加
            If aPresetFilter(iRow).intUnSharpLengthUp > 0 Then
                Call SubImageFiltering("miUnSharpLengthUp", f.SelectedImage, aPresetFilter(iRow).intUnSharpLengthUp)
            End If
            
            '图像增强幅度减少
            If aPresetFilter(iRow).intUnSharpLengthDown > 0 Then
                Call SubImageFiltering("miUnSharpLengthDown", f.SelectedImage, aPresetFilter(iRow).intUnSharpLengthDown)
            End If
            
            '平滑增加
            If aPresetFilter(iRow).intFilterLengthUp > 0 Then
                Call SubImageFiltering("miFilterLengthUp", f.SelectedImage, aPresetFilter(iRow).intFilterLengthUp)
            End If
            
            '平滑减少
            If aPresetFilter(iRow).intFilterLengthDown > 0 Then
                Call SubImageFiltering("miFilterLengthDown", f.SelectedImage, aPresetFilter(iRow).intFilterLengthDown)
            End If
        End If
    End If
    
    '处理序列内图像同步
    Call subSeriesInPhase(f.intSelectedSerial, f, f.SelectedImage, IMG_SYN_FILTER)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subFunctionWL(ByVal control As CommandBarControl, f As Form)
'------------------------------------------------
'功能：功能键设置窗宽窗位处理
'参数： Control--菜单控件；
'       f--窗体。
'返回：无
'------------------------------------------------
    Dim i As Integer
    Dim iWWidth As Integer
    Dim iWLevel As Integer
    Dim intFormType As Integer  '1是主观片窗体，2是胶片打印窗体,3是胶片预览大窗口
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If f.SelectedImage Is Nothing Then Exit Sub
    If control Is Nothing Then Exit Sub
    
    
    If f.Name = "frmFilmView" Then
        intFormType = 3
    ElseIf f.Name = "frmFilm" Then
        intFormType = 2
    Else
        intFormType = 1
    End If
    
    If control.Id = ID_Active_AdjustWindow_HandAdjustWindow_Custom Then
        frmWindowCustom.lngWindow = f.SelectedImage.width
        frmWindowCustom.lngLevel = f.SelectedImage.Level
        frmWindowCustom.Show 1, f
        If frmWindowCustom.bApply Then
            control.Category = frmWindowCustom.lngWindow & "-" & frmWindowCustom.lngLevel
        End If
    End If
    i = InStr(control.Category, "-")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If i <> 0 Then
        f.SelectedImage.width = Val(Mid(control.Category, 1, i - 1))
        f.SelectedImage.Level = Val(Mid(control.Category, i + 1))
        f.SelectedImage.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & f.SelectedImage.width & "-L:" & f.SelectedImage.Level
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If control.Id = ID_Active_AdjustWindow_HandAdjustWindow_ReSet Then
        f.SelectedImage.VOILUT = 1
        '判断是否有两个默认窗口
        If f.blnDefaultWW2 = False Then
            '显示默认第二个窗口
            If f.SelectedImage.Attributes(&H28, &H1050).VM = 2 And f.SelectedImage.Attributes(&H28, &H1051).VM = 2 Then
                iWWidth = f.SelectedImage.Attributes(&H28, &H1051).ValueByIndex(2)
                iWLevel = f.SelectedImage.Attributes(&H28, &H1050).ValueByIndex(2)
                f.SelectedImage.width = iWWidth
                f.SelectedImage.Level = iWLevel
                f.blnDefaultWW2 = True
            Else
                f.SelectedImage.SetDefaultWindows
            End If
        Else
            f.SelectedImage.SetDefaultWindows
            f.blnDefaultWW2 = False
        End If
        
        
        If f.SelectedImage.Attributes(&H6000, &H15).Value = 1 Then
            If f.SelectedImage.Level = 0 Then
                f.SelectedImage.Level = 1
            End If
        End If
    End If
    '处理序列内图像同步
    If intFormType = 1 Then '主观片窗体
        Call subSeriesInPhase(f.intSelectedSerial, f, f.SelectedImage, IMG_SYN_WINDOW)
    ElseIf intFormType = 2 Then     '胶片预览窗体
        Call f.subSynchronalImg(False, IMG_SYN_WINDOW)
    End If
End Sub

Public Sub subcalibrate(f As frmViewer)
'------------------------------------------------
'功能：校准
'参数：f－－窗体
'返回：无
'2009用
'------------------------------------------------
    Dim va As Variant, l As DicomLabel
    If f.SelectedImage Is Nothing Or f.SelectedLabel Is Nothing Then
        MsgBox "校准必须选择一个直线标注", vbInformation, gstrSysName
        Exit Sub
    End If
    If f.SelectedLabel.LabelType <> doLabelLine Then
        MsgBox "校准必须选择一个直线标注", vbInformation, gstrSysName
        Exit Sub
    End If
    va = f.SelectedImage.Attributes(&H28, &H30).Value
    If Not IsNull(va) Then
        If MsgBox("该图像已经有校准信息,是否重新校准?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    End If
    
    Dim strResult As String
    strResult = 0
    strResult = InputBox("该直线原来长度为：" & f.SelectedLabel.ROILength, "直线校准", f.SelectedLabel.ROILength, _
                    f.left + f.width / 4, f.top + f.height / 4)
    If strResult = "" Then Exit Sub
    If strResult < 0 Then
        MsgBox "校准长度应该大于0，请重新校准。", vbInformation, gstrSysName
        Exit Sub
    Else
        f.dubCalibrateLength = Val(strResult)
    End If
    If IsNull(va) Then
        ReDim va(1 To 2)
        va(1) = f.dubCalibrateLength / f.SelectedLabel.ROILength
        va(2) = f.dubCalibrateLength / f.SelectedLabel.ROILength
    Else
        va(1) = va(1) * f.dubCalibrateLength / f.SelectedLabel.ROILength
        va(2) = va(2) * f.dubCalibrateLength / f.SelectedLabel.ROILength
    End If
    f.SelectedImage.Attributes.Add &H28, &H30, va
    For Each l In f.SelectedImage.Labels
        If f.SelectedImage.Labels.IndexOf(l) > G_INT_SYS_LABEL_COUNT And l.LabelType = doLabelText And l.Tag = "RIO" Then    '''''测量性质的文字标注应该给一个类型便于识别
            l.Text = l.TagObject.ROILength & l.TagObject.ROIDistanceUnits
        End If
    Next
    ''''''这里注意把原来的测量信息写入到图像的一个LABEL中,以便保存到图像中
    f.SelectedImage.Refresh False
End Sub

Public Sub SubImageFiltering(strFilterString As String, img As DicomImage, Optional intTimes As Integer = 1)
'------------------------------------------------
'功能：真正处理图像处理，包括边缘增强，平滑处理和复原
'参数： strFilterString--表示虑镜处理类型的字符串；
'       img--需要处理的图像
'       intTimes -- 图像处理的次数
'返回：直接对图像进行处理
'------------------------------------------------
    If img Is Nothing Then Exit Sub
    If intTimes <= 0 Then Exit Sub
    
    Dim dblUnsharpEnhancement As Double
    Dim intUnsharpLength As Integer
    Dim intFilterLength As Integer
    
    dblUnsharpEnhancement = img.UnsharpEnhancement
    intUnsharpLength = img.UnsharpLength
    intFilterLength = img.FilterLength
    
    Select Case strFilterString
    Case "miUnSharpEnhancementUp"      '边缘增强强度增加，幅度0.1
        dblUnsharpEnhancement = dblUnsharpEnhancement + intTimes * 0.1
        If dblUnsharpEnhancement < 30 Then
            img.UnsharpEnhancement = dblUnsharpEnhancement
            If img.UnsharpLength = 0 Then img.UnsharpLength = 1
        End If
    Case "miUnSharpEnhancementDown"         '边缘增强强度减少，幅度0.1
        dblUnsharpEnhancement = dblUnsharpEnhancement - intTimes * 0.1
        If dblUnsharpEnhancement >= 0 Then
            img.UnsharpEnhancement = dblUnsharpEnhancement
            If img.UnsharpLength = 0 Then img.UnsharpLength = 1
        Else
            img.UnsharpEnhancement = 0
        End If
    Case "miUnSharpLengthUp"   '边缘增强幅度增加，幅度1
        intUnsharpLength = intUnsharpLength + intTimes
        If intUnsharpLength < 30 Then
            img.UnsharpLength = intUnsharpLength
            If img.UnsharpEnhancement = 0 Then img.UnsharpEnhancement = 0.1
        End If
    Case "miUnSharpLengthDown"   '边缘增强幅度减少，幅度1
        intUnsharpLength = intUnsharpLength - intTimes
        If intUnsharpLength >= 0 Then
            img.UnsharpLength = intUnsharpLength
            If img.UnsharpEnhancement = 0 Then img.UnsharpEnhancement = 0.1
        Else
            img.UnsharpLength = 0
        End If
    Case "miFilterLengthUp"       '平滑增加，幅度1
        '判断Zoom是否＝1，如果是，则修改为0.9999
        If img.ActualZoom = 1 Then
            img.Zoom = 0.9999
        End If
        '判断图像在放大和缩小模式下面，是否是doFilterMovingAverage，只有这个模式下才可以平滑
        '缩小模式下，默认值就是doFilterMovingAverage，不用修改
        img.MagnificationMode = doFilterMovingAverage
        
        '判断FilterLength是否＝0如果是，则在2/ActualZoom和2×FilterLength之间进行调整
        If intFilterLength = 0 Then
            If intTimes = 1 Then
                img.FilterLength = 2 / img.ActualZoom + 1
            ElseIf intTimes > 1 Then
                img.FilterLength = 2 / img.ActualZoom + 1
                img.FilterLength = img.FilterLength + (intTimes - 1)
            End If
        Else    '正常情况下FilterLength＋1
            img.FilterLength = intFilterLength + intTimes
        End If
    Case "miFilterLengthDown"    '平滑减少，幅度1
        '判断Zoom是否＝1，如果是，则修改为0.9999
        If img.ActualZoom = 1 Then
            img.Zoom = 0.9999
        End If
        '判断图像在放大和缩小模式下面，是否是doFilterMovingAverage，只有这个模式下才可以平滑
        '缩小模式下，默认值就是doFilterMovingAverage，不用修改
        img.MagnificationMode = doFilterMovingAverage
        
        '判断当前FilterLength－1是否小于 2/ActualZoom
        intFilterLength = intFilterLength - intTimes
        img.FilterLength = IIf(intFilterLength < 2 / img.ActualZoom, 0, intFilterLength)
    Case "miRestore"     '图像还原
        img.UnsharpEnhancement = 0
        img.UnsharpLength = 0
        img.FilterLength = 0
    End Select
End Sub

Public Sub SubImageUnsharp(strFilterString As String, f As frmViewer)
'------------------------------------------------
'功能：图像处理，包括边缘增强，平滑处理和复原
'参数：strFilterString--表示虑镜处理类型的字符串；f--图像增强的窗体
'返回：无
'2009用
'------------------------------------------------
    If f.SelectedImage Is Nothing Then Exit Sub
    
    Call SubImageFiltering(strFilterString, f.SelectedImage)
    
    f.Viewer(f.intSelectedSerial).Refresh
    ''''''''''''''''''''''''''''''''''''''''作序列内图像同步'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call subSeriesInPhase(f.intSelectedSerial, f, f.SelectedImage, IMG_SYN_FILTER)
End Sub

Public Sub subSetImageFortF(f As frmViewer)
'------------------------------------------------
'功能：根据当前选中序列的排序方式，设置排序菜单的选择状态
'参数：f--排序的窗体
'返回：无
'------------------------------------------------
    Dim iSortType As Integer    '记录当前序列的排序方式：0--图像号；1--床位正序；2--床位逆序；3--采集时间；4--图像时间，仅在ZLShowSeriesInfos中使用。
    Dim lngToolID As Long
    
    On Error GoTo err
    
    If f.intSelectedSerial = 0 Then Exit Sub
    iSortType = ZLShowSeriesInfos(f.intSelectedSerial).intSortType
    
    Select Case iSortType
            Case 0            '按照图像号排序
                lngToolID = ID_View_PhotoSerial_PhotoNumber
            Case 1                 '按照床位正序排序
                lngToolID = ID_View_PhotoSerial_BedASC
            Case 2                '按照床位逆序排序
                lngToolID = ID_View_PhotoSerial_BedDESC
            Case 3         '按照采集时间排序
                lngToolID = ID_View_PhotoSerial_CollectionTime
            Case 4              '按照图像时间排序
                lngToolID = ID_View_PhotoSerial_PhotoTime
        End Select
        
    '处理菜单和工具条的check状态
    f.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_PhotoSerial_PhotoNumber, , True).Checked = False
    f.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_PhotoSerial_BedASC, , True).Checked = False
    f.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_PhotoSerial_BedDESC, , True).Checked = False
    f.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_PhotoSerial_CollectionTime, , True).Checked = False
    f.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_PhotoSerial_PhotoTime, , True).Checked = False

    f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_PhotoSerial_PhotoNumber, , True).Checked = False
    f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_PhotoSerial_BedASC, , True).Checked = False
    f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_PhotoSerial_BedDESC, , True).Checked = False
    f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_PhotoSerial_CollectionTime, , True).Checked = False
    f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_PhotoSerial_PhotoTime, , True).Checked = False

    '选中排序方式
    f.ComToolBar.Item(ToolBar_Comm).FindControl(, lngToolID, , True).Checked = True
    f.ComToolBar.Item(ToolBar_Menu).FindControl(, lngToolID, , True).Checked = True
    f.ComToolBar.RecalcLayout
        
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub subMnuImageSort(lngToolID As Long, f As frmViewer)
'------------------------------------------------
'功能：排序方式处理，根据lngToolID来处理排序，同时处理菜单的选择状态
'参数：Tool--进行排序的工具栏控件；f--排序的窗体
'返回：无
'2009用
'------------------------------------------------
    Dim iSortType As Integer        '0--图像号；1--床位正序；2--床位逆序；3--采集时间；4--图像时间。
    
    '清除所有的排序方式
    If lngToolID = ID_View_PhotoSerial_PhotoNumber Or lngToolID = ID_View_PhotoSerial_BedASC _
        Or lngToolID = ID_View_PhotoSerial_BedDESC Or lngToolID = ID_View_PhotoSerial_CollectionTime _
        Or lngToolID = ID_View_PhotoSerial_PhotoTime Then
        
        Select Case lngToolID
            Case ID_View_PhotoSerial_PhotoNumber            '按照图像号排序
                iSortType = 0
            Case ID_View_PhotoSerial_BedASC                 '按照床位正序排序
                iSortType = 1
            Case ID_View_PhotoSerial_BedDESC                '按照床位逆序排序
                iSortType = 2
            Case ID_View_PhotoSerial_CollectionTime         '按照采集时间排序
                iSortType = 3
            Case ID_View_PhotoSerial_PhotoTime              '按照图像时间排序
                iSortType = 4
        End Select
        
        'intSelectedSerial存在，而且被选中的序列中有图像，才进行排序
        If f.intSelectedSerial > 0 And f.intSelectedSerial < f.MSFViewer.Rows And f.MSFViewer.TextMatrix(f.intSelectedSerial, 1) = "True" Then
            
            Call subSortImages(f, f.intSelectedSerial, iSortType)
            '强制让滚动条刷新一次
            Call subShowALLImage(f, f.Viewer(f.intSelectedSerial), 1, False)
            f.VScro(f.intSelectedSerial).Value = 1
            
            '根据排序方式，设置菜单勾选
            Call subSetImageFortF(f)
        End If
    End If
End Sub

Public Sub subMouseRLset(ByVal control As CommandBarControl)
'------------------------------------------------
'功能：处理鼠标左右键的check状态。
'参数：Control--工具栏控件
'返回：无
'2009用
'------------------------------------------------
    Dim i As Integer, j As Integer
    For i = 1 To cMouseUsage.Count
        If cMouseUsage(i).ButtomID = control.Id Then Exit For
    Next
    If i <= cMouseUsage.Count Then
        For j = 1 To cMouseUsage.Count
            If cMouseUsage(j).strProgramName <> "No" Then
                If cMouseUsage(j).ButtomID <> control.Id And cMouseUsage(i).lngMouseKey = cMouseUsage(j).lngMouseKey And cMouseUsage(i).lngShift = cMouseUsage(j).lngShift Then
                    control.Checked = False
                Else
                    control.Checked = True
                End If
            End If
        Next
    End If
End Sub

Public Sub subCurrentCheck(control As CommandBarControl, f As frmViewer)
'------------------------------------------------
'功能：定位线处理，处理定位线菜单的单击事件，控制定位线三个相关按钮只能被选中一个，按钮状态保存到临时变量。
'参数：Control--被单击的菜单；f--显示定位线的窗体。
'返回：无
'2009用
'------------------------------------------------
    If control.Id = ID_Active_PointingLine_ALL Then
        f.ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_FirstLast, , True).Checked = False
        f.ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_Now, , True).Checked = False
    End If
    If control.Id = ID_Active_PointingLine_FirstLast Then
        f.ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_ALL, , True).Checked = False
    End If
    If control.Id = ID_Active_PointingLine_Now Then
        f.ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_ALL, , True).Checked = False
    End If
    f.ComToolBar.Item(ToolBar_Plane).FindControl(, control.Id, , True).Checked = Not control.Checked
    
    Button_miAllReferLine = f.ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_ALL).Checked
    Button_miFLReferLine = f.ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_FirstLast).Checked
    Button_miCurrentReferLine = f.ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_Now).Checked
    
    '如果进入了胶片打印状态，则在胶片打印窗体中也显示定位线
    If f.blnPrintFilm = True And Not f.mfrmFilm Is Nothing Then
        f.mfrmFilm.subDispReferLineFilm
    End If
    
    '显示定位线
    Call subDisplayReferLine(f.Viewer(f.intSelectedSerial), f, False)
End Sub

Public Sub subOutputToPowerPoint(f As frmViewer)
'------------------------------------------------
'功能：输出到POWERPOINT
'参数：f--窗体
'返回：无
'------------------------------------------------
    Dim v As DicomViewer
    Dim im As DicomImage
    Dim imgs As New DicomImages
    Dim iW, iH As Integer               '原始图像的宽和高
    Dim ix, iy As Integer               '现在的起点坐标
    Dim Nw, Nh As Integer               '现在的宽和高
    Dim intCol, intRow As Integer       '当前列数和行数
    Dim NowImg As Integer               '当前的图片
    Dim ImgCount As Integer             '可显示的图像数
    Dim ShowImg As Integer              '可显示的图像的起始位置
    Dim j As Integer                    '循环变量
    Dim z As Integer                    '临时变量
    Dim x As Integer                    '循环变量
    Dim i As Integer                    '循环变量
    Dim JSCount As Integer              '记录位置
    Dim TwoBegin                        '第二个开始位置
    Dim PageCount As Integer            '一页的总数量
    Dim ppt As Object                   'PowerPoint对像
    Dim blnHaveImage As Boolean
    
    '打开PowerPoint
    Set ppt = CreateObject("PowerPoint.Application")
    
    '判断是否有图像
    For Each v In f.Viewer
        If v.Index <> 0 And v.Visible Then
            For Each im In v.Images
                If im.Tag <> "" Then
                    blnHaveImage = True
                    Exit For
                End If
            Next
            If blnHaveImage = True Then Exit For
        End If
    Next
    
    If blnHaveImage = False Then
        MsgBox "当前没有选择任何图像,不能输出!", vbInformation, gstrSysName
        Exit Sub
    End If
    '初使化
    ppt.Visible = True
    ppt.Presentations.Add 1
    ppt.ActiveWindow.view.GotoSlide (ppt.ActivePresentation.Slides.Add(1, 12).SlideIndex)
    
    '初使化位置为1
    JSCount = 1
    
    For x = 1 To f.Viewer.Count - 1
        If f.Viewer(x).Index <> 0 And f.Viewer(x).Visible Then
            imgs.Clear
            '写入图像
            For Each im In f.Viewer(x).Images
                If im.Tag <> "" Then imgs.Add im
            Next
            If imgs.Count <> 0 Then
                PageCount = f.Viewer(x).MultiColumns * f.Viewer(x).MultiRows
                j = 1
                For Each im In imgs
                    If j > PageCount Then
                        z = ppt.ActivePresentation.Slides.Count
                        ppt.ActiveWindow.view.GotoSlide (ppt.ActivePresentation.Slides.Add(z + 1, 12).SlideIndex)
                        j = 1
                    End If
                    im.Copy
                    ppt.ActiveWindow.view.Paste
                    j = j + 1
                Next
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                For i = JSCount To ppt.ActivePresentation.Slides.Count
                    ppt.ActiveWindow.view.GotoSlide i
                    If ppt.ActiveWindow.Selection.SlideRange.Shapes.Count > 0 Then
                        With ppt.ActiveWindow.Selection.SlideRange
                            ix = .Shapes(1).left
                            iy = .Shapes(1).top
                            iW = .Shapes(1).width / f.Viewer(x).MultiColumns
                            iH = .Shapes(1).height / f.Viewer(x).MultiRows
                        End With
                        For j = 1 To ppt.ActiveWindow.Selection.SlideRange.Shapes.Count
                            '得到当前图像位置
                            If (j Mod f.Viewer(x).MultiColumns) = 0 Then
                                intRow = j / f.Viewer(x).MultiColumns
                                intCol = f.Viewer(x).MultiColumns
                            Else
                                intRow = Int(j / f.Viewer(x).MultiColumns) + 1
                                intCol = j Mod f.Viewer(x).MultiColumns
                            End If
                            '移动图像位置
                            With ppt.ActiveWindow.Selection.SlideRange
                                .Shapes(j).top = iy + (iH * (intRow - 1))
                                .Shapes(j).left = ix + (iW * (intCol - 1))
                                If iH > iW Then
                                    .Shapes(j).width = iW
                                Else
                                    .Shapes(j).height = iH
                                End If
                            End With
                        Next
                        JSCount = JSCount + 1
                    End If
                Next
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '最后一行时不再增加页
                If x < f.Viewer.Count - 1 Then
                    z = ppt.ActivePresentation.Slides.Count
                    ppt.ActiveWindow.view.GotoSlide (ppt.ActivePresentation.Slides.Add(z + 1, 12).SlideIndex)
                End If
            End If
        End If
    Next
    ppt.ActiveWindow.view.GotoSlide 1
End Sub

Public Sub subDSA(thisForm As frmViewer)
'------------------------------------------------
'功能：DSA数字减影
'参数：thisForm--进行数字减影的窗体。
'返回：无
'2009用
'------------------------------------------------
    If thisForm.SelectedImage Is Nothing Then Exit Sub             ''''当前没有选择图像
    If thisForm.SelectedImage.FrameCount <= 1 Then Exit Sub        ''''当前图像不是多祯
    If IsNull(thisForm.SelectedImage.Attributes(&H28, &H4).Value) Then Exit Sub
    If Mid(thisForm.SelectedImage.Attributes(&H28, &H4).Value, 1, 4) <> "MONO" Then Exit Sub
    
    Call FrmDSAConfig.zlShowMe(thisForm.SelectedImage.FrameCount, thisForm.SelectedImage.Frame, thisForm)
End Sub

Public Sub subCutOut(f As frmViewer)
'------------------------------------------------
'功能：进入和退出裁减状态，隐藏或显示裁剪标注
'参数：f--进入和退出裁剪状态的窗体
'返回：无，直接控制裁剪状态标注的显示和隐藏
'------------------------------------------------
    Dim i As Integer, im As DicomImage
    If f.SelectedImage Is Nothing Then Exit Sub
    ''''''''''''[退出裁减]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not Button_miCutOut Then
        For i = 1 To 5
            f.SelectedImage.Labels(i).Visible = False '
            f.SelectedImage.Labels(i).Tag = f.SelectedImage.Labels(i).left & "_" & f.SelectedImage.Labels(i).top    ''''记录原始状态，供再次显示用
            f.SelectedImage.Labels(i).left = G_INT_SYS_LABEL_HIDE_LEFT
            f.SelectedImage.Labels(i).top = G_INT_SYS_LABEL_HIDE_TOP
        Next
        If Not f.SelectedLabel Is Nothing Then
            If f.SelectedImage.Labels.IndexOf(f.SelectedLabel) = 1 Then          ''''如果当前选择的是裁减标注，则取消显示句柄
                SubNoDispPeriod f.SelectedImage, f      '为指定图像隐藏标注选择句柄
                Set f.SelectedLabel = Nothing
            End If
        End If
    Else
        ''''''''''''[进入裁减]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If f.SelectedImage.Labels(1).Tag <> "" Then         '''''如果以前裁减过则回复以前的显示位置
            For i = 1 To 5
                f.SelectedImage.Labels(i).left = Val(left(f.SelectedImage.Labels(i).Tag, InStr(f.SelectedImage.Labels(i).Tag, "_") - 1))
                f.SelectedImage.Labels(i).top = Val(Right(f.SelectedImage.Labels(i).Tag, Len(f.SelectedImage.Labels(i).Tag) - InStr(f.SelectedImage.Labels(i).Tag, "_")))
            Next
        Else
            f.SelectedImage.Labels(1).left = 4
            f.SelectedImage.Labels(1).top = 4
            f.SelectedImage.Labels(1).width = f.SelectedImage.sizex - 8
            f.SelectedImage.Labels(1).height = f.SelectedImage.sizey - 8
        End If
        SubDispPeriod f.SelectedImage.Labels(1), f.SelectedImage, f '为指定图像中的指定标注，显示标注选择句柄
        For i = 1 To 5
            f.SelectedImage.Labels(i).Visible = True
        Next
        Set f.SelectedLabel = f.SelectedImage.Labels(1)
    End If
    
    ''''''''''''''在裁减状态下对裁减操作作出图像同步处理'''''''''''''''
    If Button_miImageInPhase = True Then subCutOutInphase f.Viewer(f.intSelectedSerial), f.SelectedImage, f
    f.SelectedImage.Refresh False
    f.ComToolBar.RecalcLayout
End Sub

Public Sub subDispLabelInfo(f As frmViewer)
'------------------------------------------------
'功能：显示或隐藏图像的用户标注信息
'参数： f--需要显示或隐藏用户标注信息的窗体；
'返回：无，直接显示或隐藏用户标注。
'2009用
'------------------------------------------------
    Dim img As DicomImage
    Dim v As DicomViewer
    Dim l As DicomLabel
    Dim i As Integer
    Dim CmdControl As CommandBarControl
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If f.SelectedImage Is Nothing Then Exit Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Button_miDispLabelInfo = Not Button_miDispLabelInfo
    f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_LableShow, , True).Checked = Button_miDispLabelInfo
    f.ComToolBar.RecalcLayout
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each v In f.Viewer
        If v.Index <> 0 Then
            For Each img In v.Images
                If img.Labels.Count > G_INT_SYS_LABEL_COUNT Then
                    For i = G_INT_SYS_LABEL_COUNT + 1 To img.Labels.Count
                        If Button_miDispLabelInfo Then
                            img.Labels(i).Visible = True
                        Else
                            img.Labels(i).Visible = False
                        End If
                    Next i
                    If Not Button_miDispLabelInfo Then
                        For i = 11 To 18                    '隐藏标注句柄
                            img.Labels(i).Visible = False
                        Next i
                    End If
                End If
            Next
        End If
        v.Refresh
    Next
End Sub

Public Sub subManipulation(strOperation As String, f As frmViewer)
'------------------------------------------------
'功能：调用多平面处理（图像旋转、反白等）
'参数： strOperation--表示翻转方式的字符串； f--窗体。
'返回：无
'2009用
'------------------------------------------------
    If f.SelectedImage Is Nothing Then Exit Sub
    
    On Error GoTo err
    
    Call subFlipRotate(f.SelectedImage, strOperation)
    
    ''''''''''''''''''''序列内图像同步'''''''''''''''''''''''''''
    Select Case strOperation
    Case "FlipHorizontal", "FlipVertical"
        Call subSeriesInPhase(f.intSelectedSerial, f, f.SelectedImage, IMG_SYN_FLIP)
    Case "RotateAnticlockwise", "RotateClockwise"
        Call subSeriesInPhase(f.intSelectedSerial, f, f.SelectedImage, IMG_SYN_ROTATE)
    Case "Invert"
        Call subSeriesInPhase(f.intSelectedSerial, f, f.SelectedImage, IMG_SYN_WINDOW)
    End Select
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub subFlipRotate(img As DicomImage, strOperation As String)
'------------------------------------------------
'功能：多平面处理（图像旋转、反白等）
'参数：img－－进行处理的图像；
'      strOperation --表示翻转方式的字符串
'返回：无
'2009用
'------------------------------------------------
    With img
        Select Case strOperation
        Case "FlipHorizontal"
            .FlipState = .FlipState Xor 1
            If .RotateState = doRotateLeft Or .RotateState = doRotateRight Then
                .RotateState = (.RotateState + 2) Mod 4
            End If
        Case "FlipVertical"
            .FlipState = .FlipState Xor 2
            If .RotateState = doRotateLeft Or .RotateState = doRotateRight Then
                .RotateState = (.RotateState + 2) Mod 4
            End If
        Case "RotateAnticlockwise"
            .RotateState = (.RotateState + 1) And 3
        Case "RotateClockwise"
            .RotateState = (.RotateState + 3) And 3
        Case "Invert"
            If .VOILUT = 1 Then .VOILUT = 0
            .width = -.width
        End Select
    End With
End Sub

Public Sub subSelectAllSerial(f As frmViewer)
'------------------------------------------------
'功能：选择所有序列
'参数：f--选择序列的窗体
'返回：无
'2009用
'------------------------------------------------
    Dim v As DicomViewer
    f.isSelectAllSerial = Not f.isSelectAllSerial
    
    For Each v In f.Viewer
        If v.Visible = True Then
            ZLShowSeriesInfos(v.Index).Selected = IIf(f.isSelectAllSerial, True, False)
            subDispframe f, v
            v.Refresh
        End If
    Next
End Sub
 
Public Sub subSelectAllIMage(f As frmViewer)
'------------------------------------------------
'功能：选择所有图像
'参数：f--选择图像的窗体
'返回：无
'2009用
'------------------------------------------------
    Dim v As DicomViewer
    Dim i As Integer
    
    f.isSelectAllImage = Not f.isSelectAllImage
    
    For Each v In f.Viewer
        If v.Visible = True Then
            If v.Images.Count > 0 And (v.Index = f.intSelectedSerial Or ZLShowSeriesInfos(v.Index).Selected = True) Then
                For i = 1 To ZLShowSeriesInfos(v.Index).ImageInfos.Count
                    ZLShowSeriesInfos(v.Index).ImageInfos(i).blnSelected = IIf(f.isSelectAllImage, True, False)
                Next i
                subDispframe f, v
                v.Refresh
            End If
        End If
    Next
End Sub

Public Sub subFullScreen(Frm As frmViewer)
'------------------------------------------------
'功能：切换屏幕的全屏状态
'参数：Frm--进行全屏切换的窗体
'返回：无
'2009用
'------------------------------------------------
    Dim CmdControl As CommandBar
    Dim ToolBarTop As Long
    Dim ToolBarLeft As Long
    Dim ToolBarHeight As Long
    Dim ToolBarWidth As Long
    Dim i As Integer
    blfrmRefresh = False
    '''''''''''''''''''''''''''''''''''''''[不是全屏状态的处理]''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not Button_miFullScreen Then
        Frm.WindowState = vbMaximized
        Frm.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_FullScreen, , True).Checked = True
        Frm.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_FullScreen, , True).Checked = True
        Frm.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_FullScreen, , True).ToolTipText = "取消全屏显示"
        Frm.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_FullScreen, , True).ToolTipText = "取消全屏显示"
        Button_miFullScreen = True
        '''''''''''''''''''''''''''''''''''''''[关闭窗体的标题栏]'''''''''''''''''''''''''''''''''''''''''''''
        
        ''''设置窗体是否显示标题栏
        Call zlcontrol.FormSetCaption(Frm, False)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '隐藏工具栏和菜单
        For i = 1 To 8
            If i <> 7 Then
                Frm.ComToolBar.Item(i).Visible = False
            End If
        Next
        '隐藏状态栏
        blfrmRefresh = True
        Frm.sbStatusBar.Visible = False
        
        Set CmdControl = Frm.ComToolBar.Item(7)
        CmdControl.GetWindowRect ToolBarLeft, ToolBarTop, ToolBarWidth, ToolBarHeight
        
        Frm.ComToolBar.DockToolBar CmdControl, Frm.left, Frm.height + Frm.top - Frm.sbStatusBar.height - (ToolBarHeight - ToolBarTop), xtpBarFloating
        Frm.ComToolBar.RecalcLayout
    Else
        ''''''''''''''''''''''''''''''''''''''[回复正常屏幕]''''''''''''''''''''''''''''''''''''''''''''''
        Frm.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_FullScreen, , True).Checked = False
        Frm.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_FullScreen, , True).Checked = False
        Frm.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_FullScreen, , True).ToolTipText = "全屏显示"
        Frm.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_FullScreen, , True).ToolTipText = "全屏显示"
        Button_miFullScreen = False
        '''''''''''''''''''''''''''''''''''''''[关闭窗体的标题栏]'''''''''''''''''''''''''''''''''''''''''''''
        ''''设置窗体是否显示标题栏
        Call zlcontrol.FormSetCaption(Frm, True)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        ''''''''''''''''''''''''''''''''''''''[开启窗体的标题栏]''''''''''''''''''''''''''''''''''''''''''''''
        '显示工具栏和菜单
        For i = 1 To 8
            Frm.ComToolBar.Item(i).Visible = True
        Next
        '显示状态栏
        Frm.sbStatusBar.Visible = True
        Frm.ComToolBar.Item(7).Position = Frm.ComToolBar.Item(2).Position
        Frm.ComToolBar.RecalcLayout
        '重新按一定的顺序摆放工具栏位置
        ArrayToolBar Frm.ComToolBar, Frm.top, Frm.left, Frm.height, Frm.width
        blfrmRefresh = True
    End If
End Sub

Public Function subSaveImage(img As DicomImage, strOldSeriesUID As String) As Boolean
'------------------------------------------------
'功能：将图像保存到数据库中
'参数：img 保存的图像
'返回：True---正确保存；False ---出现错误，没有保存图像
'------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    
    Dim dtReceived As String
    Dim strStudyUID As String
    Dim blnFirstImage As String     '是否本次检查的第一张图像
    Dim lngResult As String         'FTP操作结果
    Dim NowTime As Date
    Dim strSQL As String
    
    Dim strFTPDir As String
    Dim strFTPIp As String
    Dim strFTPUser As String
    Dim strFTPPassw As String
    Dim Inet As New clsFtp             'FTP类
    
    Dim arrSQL() As Variant         '事务中的SQL语句数组
    Dim blnInTrans As Boolean       '是否正在事务处理的过程中
    Dim i As Integer
    
    subSaveImage = False
    
    If img Is Nothing Then
        MsgBox "拼接的结果图出错，无法保存。", vbOKOnly, "提示信息"
        Exit Function
    End If
    
    '先判断图象是否已经存在
    
    strSQL = "Select 图像UID From 影像检查图象 Where 图像UID=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查图象是否存在", CStr(img.InstanceUID))
    If rsTmp.EOF = False Then
        MsgBox "数据库中找不到此图像，无法保存外部直接打开的临时图像。", vbOKOnly, "提示信息"
        Exit Function
    End If
    
    '先保存FTP图像
    '读取接收日期
    strSQL = "select a.接收日期,a.检查UID  from 影像检查记录 a,影像检查序列 b where a.检查UID =b.检查UID and b.序列UID =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取检查UID", strOldSeriesUID)
    
    If rsTmp.EOF = True Then
        MsgBox "数据库中找不到此序列，无法保存外部直接打开的临时图像。", vbOKOnly, "提示信息"
        Exit Function '查询不到记录，则退出保存
    End If
    
    NowTime = zlDatabase.Currentdate
    strStudyUID = rsTmp("检查UID")
    dtReceived = Format(rsTmp("接收日期"), "yyyyMMdd")
     
    '保存图像到缓存目录
    MkLocalDir PstrBufferImagePath & dtReceived & "/" & strStudyUID & "/"
    '不处理，保持原图的压缩方式
    img.WriteFile PstrBufferImagePath & dtReceived & "/" & strStudyUID & "/" & img.InstanceUID, True
    
    '连接FTP
    Call funGetStorageDevice(strStudyUID, strFTPDir, strFTPIp, strFTPUser, strFTPPassw)
    lngResult = Inet.FuncFtpConnect(strFTPIp, strFTPUser, strFTPPassw)
    
    '保存图像文件
    If lngResult = 0 Then
        'FTP操作失败，提示错误，并删除缩略图中的图像
        MsgBox "FTP连接失败，图像无法保存，请检查网络设置。", vbInformation, gstrSysName
        Exit Function
    Else
        '在FTP中创建目录
        Inet.FuncFtpMkDir "/", strFTPDir
        
        '向FTP上传文件
        Inet.FuncUploadFile strFTPDir, _
             PstrBufferImagePath & dtReceived & "/" & strStudyUID & "/" & img.InstanceUID, img.InstanceUID
    End If
    Inet.FuncFtpDisConnect
    
    '图像存储成功后，存储数据库信息
    On Error GoTo DBError
    arrSQL = Array()
    
    strSQL = "Select 序列UID From 影像检查序列  Where 序列UID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PACS图像保存", CStr(img.SeriesUID))
    '插入新的检查序列
    If rsTmp.EOF Then
        strSQL = "ZL_影像序列_INSERT('" & strStudyUID & "','" & img.SeriesUID & "','" & _
            img.SeriesDescription & "',0)"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
    End If
    
    '插入新的图像
    strSQL = "ZL_影像图象_INSERT('" & img.InstanceUID & "','" & img.SeriesUID & "','" & _
        img.SeriesDescription & "',0," & IIf(GetImageAttribute(img.Attributes, ATTR_图像号) = "", 0, GetImageAttribute(img.Attributes, ATTR_图像号)) & ","
    If GetImageAttribute(img.Attributes, ATTR_采集日期) <> "" And GetImageAttribute(img.Attributes, ATTR_采集时间) <> "" Then
        strSQL = strSQL & "to_Date('" & Format(GetImageAttribute(img.Attributes, ATTR_采集日期) & " " & GetImageAttribute(img.Attributes, ATTR_采集时间), "yyyy-MM-dd HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS'),"
    Else
        strSQL = strSQL & " sysdate,"
    End If
    
    If GetImageAttribute(img.Attributes, ATTR_图像日期) <> "" And GetImageAttribute(img.Attributes, ATTR_图像时间) <> "" Then
        strSQL = strSQL & "to_Date('" & Format(GetImageAttribute(img.Attributes, ATTR_图像日期) & " " & GetImageAttribute(img.Attributes, ATTR_图像时间), "yyyy-MM-dd HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS'),'"
    Else
        strSQL = strSQL & " sysdate,'"
    End If
    
        strSQL = strSQL & GetImageAttribute(img.Attributes, ATTR_层厚) & "','" _
        & GetImageAttribute(img.Attributes, ATTR_图像位置病人) & "','" _
        & GetImageAttribute(img.Attributes, ATTR_图像方向病人) & "','" _
        & GetImageAttribute(img.Attributes, ATTR_参考帧UID) & "','" _
        & GetImageAttribute(img.Attributes, ATTR_切片位置) & "','" _
        & GetImageAttribute(img.Attributes, ATTR_行数) & "','" _
        & GetImageAttribute(img.Attributes, ATTR_列数) & "','" _
        & GetImageAttribute(img.Attributes, ATTR_像素距离) & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
    
    '开始事务处理
    gcnOracle.BeginTrans
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "保存图像")
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False
    
    subSaveImage = True
    
    Exit Function
DBError:
    If blnInTrans = True Then gcnOracle.RollbackTrans
    Inet.FuncFtpDisConnect
    err.Raise err.Number, "检查图像保存"
End Function


