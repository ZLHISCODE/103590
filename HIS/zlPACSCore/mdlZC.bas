Attribute VB_Name = "mdlPopupMenu"
Option Explicit
'--------------------------------------------------------
'功  能：本模块根据当前的图像情况产生图像复制用弹出菜单（需要优化为增加图像时产生菜单，需要时候调用）
'编制人：曾超
'编制日期：2004.6.12
'过程函数清单：
'    BRName():       从图像提取病人姓名
'    BRUID():        从图像提取病人ID
'    CKUID():        从图像提取检查UID
'    CheckDate():        从图像提取日期
'    CheckTime():        从图像提取时间
'    SeriesNum():        从图像提取序列号
'    CheckPart():        从图像提取部位
'    CheckMenuClass():   检查数线菜单里是否重复
'    PopMenu():      生成右键弹出的打开图像菜单
'    FiltrateStr():      过滤字串
'修改记录：
'    2005.07.08    黄捷
'-------------------------------------------------------


'从图像提取病人姓名
Function BRName(Image As DicomImage) As String
    If IsNull(Image.Attributes(&H10, &H10)) = False Then
        BRName = "姓名：" & Image.Attributes(&H10, &H10)
    End If
End Function

'从图像提取图像类型
Function IMGModality(Image As DicomImage) As String
    If IsNull(Image.Attributes(&H8, &H60)) = False Then
        IMGModality = "影像类别：" & Image.Attributes(&H8, &H60)
    End If
End Function

'从图像提取病人ID
Function BRUID(Image As DicomImage) As String
    If IsNull(Image.Attributes(&H10, &H20)) = False Then
        BRUID = Image.Attributes(&H10, &H20)
    End If
End Function

'从图像提取检查UID
Function CKUID(Image As DicomImage) As String
    If IsNull(Image.Attributes(&H20, &HD)) = False Then
        CKUID = Image.Attributes(&H20, &HD)
    End If
End Function

'从图像提取日期
Function CheckDate(Image As DicomImage) As String
    With Image
        If IsNull(.Attributes(&H8, &H22)) = False Then
            CheckDate = "日期：" & .Attributes(&H8, &H22)
        End If
        If Len(Trim(CheckDate)) < 1 Then
            If IsNull(.Attributes(&H8, &H23)) = False Then
                CheckDate = "日期：" & .Attributes(&H8, &H23)
            End If
        End If
    End With
End Function

'从图像提取时间
Function CheckTime(Image As DicomImage) As String
    On Error Resume Next
    With Image
        If IsNull(.Attributes(&H8, &H30)) = False Then
            CheckTime = "时间：" & .Attributes(&H8, &H30)
        End If
        If Len(Trim(CheckTime)) < 1 Then
            If IsNull(.Attributes(&H8, &H32)) = False Then
                CheckTime = "时间：" & .Attributes(&H8, &H32)
            End If
        End If
        If Len(Trim(CheckTime)) < 1 Then
            If IsNull(.Attributes(&H8, &H33)) = False Then
                CheckTime = "时间：" & .Attributes(&H8, &H33)
            End If
        End If
    End With
End Function

'从图像提取序列号
Function SeriesNum(Image As DicomImage) As String
    With Image
        If IsNull(.Attributes(&H20, &H11)) = False Then
            SeriesNum = "序列号：" & .Attributes(&H20, &H11)
        End If
    End With
End Function

'从图像提取部位
Function CheckPart(Image As DicomImage) As String
    With Image
        If IsNull(Image.Attributes(&H18, &H15)) = False Then
            CheckPart = "部位：" & Image.Attributes(&H18, &H15)
        End If
    End With
End Function

'检查数线菜单里是否重复
Function CheckMenuClass(CheckStr As Variant, MenuName As String, intLevel As Integer, intOneLevel As Integer, intTwoLevel As Integer) As Boolean
    '---------------------------------------------------------------------------------
    '功能：                                  搜索一次本级是否有重复的字串
    '参数：
    '           CheckStr                     数组字串
    '           MenuName                     要检查是否重复的字串
    '           intLevel                     当前级数
    '           intOneLevel                  第一级中第N个
    '           intTwoLevel                  第二级中第N个
    '返回：                                  =True表示有重复  =Flase表示没有重复
    '上级函数或过程：                        无
    '下级函数或过程：                        无
    '引用的外部参数：                        无
    '编制人：                                曾超 2005-7-7
    '----------------------------------------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim z As Integer
    Select Case intLevel
        Case 1
            For i = 0 To 20
                If CheckStr(i, 0, 0) = MenuName Then
                    CheckMenuClass = True
                    Exit Function
                End If
            Next
        Case 2
            For j = 0 To 40
                If CheckStr(intOneLevel - 1, j, 0) = MenuName Then
                    CheckMenuClass = True
                    Exit Function
                End If
            Next
        Case 3
            For z = 0 To 60
                If CheckStr(intOneLevel - 1, intTwoLevel, z) = MenuName Then
                    CheckMenuClass = True
                    Exit Function
                End If
            Next
    End Select
    CheckMenuClass = False
End Function

Public Sub PopMenu(f As frmViewer, imgs As DicomImages)
'------------------------------------------------
'功能：生成右键弹出的打开图像菜单
'参数：f--显示弹出菜单的窗体；imgs--生成弹出菜单的图像，这些图像是每个序列的第一幅图像。
'返回：无
'上级函数或过程：frmViewer.picViewer_MouseUp
'下级函数或过程：
'引用的外部参数：
'编制人：
'------------------------------------------------
    '右键菜单
    '定义菜单数组
    Dim MenuClass(20, 40, 60) As String
    Dim MenuTag(20, 40, 60) As Integer
    Dim UserName As String
    Dim CheckName As String
    Dim CKTime As String
    Dim CheckCKTime As String
    Dim CKPart As String
    '定义循环变量
    Dim i As Integer
    Dim j As Integer
    Dim z As Integer
    Dim k As Integer
    Dim l As Integer
    Dim OneClass As Integer
    Dim TwoClass As Integer
    
    Dim PopupBar As CommandBar
    Dim ControlClass1 As CommandBarPopup
    Dim ControlClass2 As CommandBarPopup
    Dim ControlClass3 As CommandBarControl
    
    If imgs.Count < 1 Then
        Exit Sub
    End If
    
    '********************姓名**************************
    For i = 1 To imgs.Count
        UserName = BRUID(imgs(i)) & "," & BRName(imgs(i)) & "," & IMGModality(imgs(i))
        If Len(Trim(UserName)) > 1 Then
            If CheckMenuClass(MenuClass, UserName, 1, 1, 1) = False Then
                MenuClass(OneClass, 0, 0) = UserName
                '记录Viewer
                MenuTag(OneClass, 0, 0) = i
                OneClass = OneClass + 1
            End If
        End If
    Next
    '***************************************************
    '*******************检查时间************************
    For i = 1 To OneClass
        k = 1
        UserName = MenuClass(i - 1, 0, 0)
        For j = 1 To imgs.Count
            CheckName = BRUID(imgs(j)) & "," & BRName(imgs(j)) & "," & IMGModality(imgs(j))
            If UserName = CheckName Then
                CKTime = CheckDate(imgs(j))
                CKTime = CKTime & " " & CheckTime(imgs(j))
                CKTime = CKUID(imgs(j)) & "," & CKTime
                CKTime = Trim(CKTime)
                If Len(CKTime) > 1 And CheckMenuClass(MenuClass, CKTime, 2, i, 1) = False Then
                    MenuClass(TwoClass, k, 0) = CKTime
                    '记录Viewer
                    MenuTag(TwoClass, k, 0) = j
                    k = k + 1
                    TwoClass = TwoClass + 1
                End If
            End If
        Next
    Next
    '********************序列+部位+序列描述************************
    For i = 1 To OneClass
        UserName = MenuClass(i - 1, 0, 0)
        For j = 1 To TwoClass
            CKTime = MenuClass(i - 1, j, 0)
            k = 1
            For z = 1 To imgs.Count
                CheckName = BRUID(imgs(z)) & "," & BRName(imgs(z)) & "," & IMGModality(imgs(z))
                CheckCKTime = CheckDate(imgs(z))
                CheckCKTime = CheckCKTime & " " & CheckTime(imgs(z))
                CheckCKTime = CKUID(imgs(z)) & "," & CheckCKTime
                If UserName = CheckName And CKTime = CheckCKTime Then
                    CKPart = "," & SeriesNum(imgs(z))
                    CKPart = CKPart & "," & CheckPart(imgs(z))
                    CKPart = CKPart & ",序列描述：" & imgs(z).SeriesDescription
                    If Len(Trim(CKPart)) > 0 And CheckMenuClass(MenuClass, CKPart, 3, i, j) = False Then
                        MenuClass(i - 1, j, k) = CKPart
                        MenuTag(i - 1, j, k) = z
                        k = k + 1
                        '记录Viewer
                        
                    End If
                End If
            Next
        Next
    Next
    
    '**********************************************************
    '新增弹出菜单
        Set PopupBar = f.ComToolBar.Add("弹出菜单", xtpBarPopup)
    '过滤多于的字符
    FiltrateStr MenuClass
    
    '生成菜单
    k = 499
    For i = 0 To OneClass - 1
        If Len(MenuClass(i, 0, 0)) > 0 And OneClass > 1 Then
            With PopupBar
                k = k + 1
                Set ControlClass1 = PopupBar.Controls.Add(xtpControlButtonPopup, k, MenuClass(i, 0, 0))
            End With
        End If
        For j = 1 To TwoClass
            If Len(MenuClass(i, j, 0)) > 0 Then
                k = k + 1
                If TwoClass > 1 Then
                    If OneClass > 1 Then
                        Set ControlClass2 = ControlClass1.CommandBar.Controls.Add(xtpControlButtonPopup, k, MenuClass(i, j, 0))
                    Else
                        Set ControlClass2 = PopupBar.Controls.Add(xtpControlButtonPopup, k, MenuClass(i, j, 0))
                    End If
                End If
                For z = 1 To 60
                    If Len(MenuClass(i, j, z)) > 0 Then
                        k = k + 1
                        If TwoClass > 1 Then
                            Set ControlClass3 = ControlClass2.CommandBar.Controls.Add(xtpControlButton, k, MenuClass(i, j, z))
                        Else
                            Set ControlClass3 = PopupBar.Controls.Add(xtpControlButton, k, MenuClass(i, j, z))
                        End If
                        ControlClass3.Category = MenuTag(i, j, z)
                    End If
                Next
            End If
        Next
    Next
    PopupBar.ShowPopup
End Sub

Private Sub FiltrateStr(MenuStr As Variant)
'------------------------------------------------
'功能：过滤字串
'参数：MenuStr--
'返回：无
'上级函数或过程：mdlPopupMenu.PopMenu
'下级函数或过程：无
'引用的外部参数：无
'编制人：曾超
'------------------------------------------------
    Dim i, j, z As Integer
    Dim StrLong As Integer
    For i = 0 To 20
        For j = 0 To 40
            For z = 0 To 60
                StrLong = InStr(MenuStr(i, j, z), ",")
                MenuStr(i, j, z) = Mid$(MenuStr(i, j, z), StrLong + 1)
            Next
        Next
    Next
End Sub

Public Sub ShowFrameSelectImagePopup(f As frmViewer) ', img As DicomImage, lblFrame As DicomLabel)
'------------------------------------------------
'功能：创建框选图象的时候 ，鼠标右键的弹出菜单
'参数：f--显示弹出菜单的窗体； img弹出菜单对应的viewer中的图像；lblFrame图象选择框
'返回：无
'------------------------------------------------

Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim cbrToolPopup As CommandBarPopup
    
    
    '鼠标右键弹出菜单
    Set cbrToolBar = f.ComToolBar.Add("鼠标右键", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ID_ACtive_SaveInReport, "保存报告图")
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub

Public Sub ShowPopup(f As frmViewer, img As DicomImage)
'------------------------------------------------
'功能：创建鼠标右键弹出菜单
'参数：f--显示弹出菜单的窗体； img弹出菜单对应的viewer中的图像，用来确定影像类别
'返回：无
'编制人：黄捷
'时间：2008-4-18
'------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim cbrToolPopup As CommandBarPopup
Dim cbrToolPopup2 As CommandBarPopup
    
    
    '鼠标右键弹出菜单
    Set cbrToolBar = f.ComToolBar.Add("鼠标右键", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ID_View_UpSeries, "上一序列")
        Set cbrControl = .Add(xtpControlButton, ID_View_DownSeries, "下一序列")
        Set cbrControl = .Add(xtpControlButton, ID_Active_Cruise, "漫游")
        Set cbrControl = .Add(xtpControlButton, ID_Active_Zoom, "缩放")
        
        Set cbrToolPopup = .Add(xtpControlButtonPopup, ID_Active_AdjustWindow_HandAdjustWindow, "手动调窗")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_AdjustWindow_HandAdjustWindow, "手动调窗")
        subSetWidthLevelF img, f, cbrToolPopup

        Set cbrControl = .Add(xtpControlButton, ID_Tool_Magnifier, "放大镜")
        Set cbrControl = .Add(xtpControlButton, ID_ACtive_FrameSelectImage, "框选图像")
        
        Set cbrToolPopup = .Add(xtpControlButtonPopup, ID_Active_Lable, "标注测量")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_ACtive_Mouse_Value, "显示CT值")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_Rect, "矩形")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_Ellipse, "椭圆")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_Area, "任意型")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_Angle, "角度")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_Arrowhead, "箭头")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_Text, "文字")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_BeeLine, "直线")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_Curve, "曲线")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_CadioThoracicRatio, "心胸比")
        
        Set cbrToolPopup = .Add(xtpControlButtonPopup, 0, "图像操作")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Eddy_LeftRight, "水平镜像")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Eddy_TopButton, "垂直镜像")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Eddy_Left90, "左转90度")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Eddy_Right90, "右转90度")
        
        Set cbrControl = .Add(xtpControlButton, ID_Tool_Movie, "电影播放")
        
        Set cbrToolPopup = .Add(xtpControlButtonPopup, ID_Active_SieveLens, "图像增强")
        
        Set cbrToolPopup2 = cbrToolPopup.CommandBar.Controls.Add(xtpControlButtonPopup, ID_Active_SieveLens_Model, "滤镜模板")
        Call subSetFilterF(img, f, cbrToolPopup2)
        
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_SieveLens_LancetAdd, "边缘增强强度增加")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_SieveLens_LancetMinus, "边缘增强强度减少")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Sievelens_LeftMoveAdd, "边缘增强幅度增加")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Sievelens_LeftMoveMinus, "边缘增强幅度减少")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_SieveLens_FlatnessAdd, "平滑增加")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_SieveLens_FlatnessMinus, "平滑减少")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Sievelens_PhotoReset, "图像还原")
        
        Set cbrToolPopup = .Add(xtpControlButtonPopup, 0, "高级图像处理")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_PointingLine_ALL, "所有定位线")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_PointingLine_FirstLast, "首尾定位线")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_PointingLine_Now, "当前定位线")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_PointingLine_3DLine, "三维鼠标")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Tool_ArrowyCoronaryReset, "MPR")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Tool_SlopeReconstruction, "斜面重建")
        
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub
