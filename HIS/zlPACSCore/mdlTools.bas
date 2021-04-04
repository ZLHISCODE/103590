Attribute VB_Name = "mdlTools"
Option Explicit
'--------------------------------------------------------
'功  能：本模块为菜单、按钮的排布等
'编制人：曾超
'编制日期：2004.6.12
'过程函数清单：
'    BarterIco():        更换图像集的Tag值。
'    CreateMenu()：      创建菜单
'    StatusBarTip():     在状态栏显示菜单的简单帮助
'    funcGetShiftStr():  通过给定shift状态数值，解析出用字符串表示的shift状态。
'    ArrayToolBar():     重新按一定的顺序摆放工具栏位置
'    ReplaceToolBarIcon():   替换当前图标为16,24,32
'    PutToolbar():       摆放工具条到Top,Left,Right,Bottom
'修改记录：
'    2005.7.02    黄捷，曾超
'-------------------------------------------------------
Public blfrmRefresh As Boolean          ''''窗体是否刷新（用于工具条刷新时窗体不一定刷新)
Public gstrSysName As String

Public Sub BarterIco(ImgBox As ImageList)
'------------------------------------------------
'功能：更换图像集的Tag值。（Tag值用来辅助实现更换工具栏图标）
'参数：ImgBox--需要更换图标Tag值的图像集。
'返回：无
'上级函数或过程：frmViewer.Form_Load；
'下级函数或过程：无
'引用的外部参数：cMouseUsage
'编制人：曾超
'------------------------------------------------
    If cMouseUsage("101").lngMouseKey = 1 Then
        ImgBox.ListImages("穿梭L").Tag = IIf(ImgBox.ListImages("穿梭L").Tag = "", ImgBox.ListImages("穿梭R").Tag, ImgBox.ListImages("穿梭L").Tag)
        ImgBox.ListImages("穿梭R").Tag = ""
    Else
        ImgBox.ListImages("穿梭R").Tag = IIf(ImgBox.ListImages("穿梭R").Tag = "", ImgBox.ListImages("穿梭L").Tag, ImgBox.ListImages("穿梭R").Tag)
        ImgBox.ListImages("穿梭L").Tag = ""
    End If
    
    If cMouseUsage("103").lngMouseKey = 1 Then
        ImgBox.ListImages("漫游L").Tag = IIf(ImgBox.ListImages("漫游L").Tag = "", ImgBox.ListImages("漫游R").Tag, ImgBox.ListImages("漫游L").Tag)
        ImgBox.ListImages("漫游R").Tag = ""
    Else
        ImgBox.ListImages("漫游R").Tag = IIf(ImgBox.ListImages("漫游R").Tag = "", ImgBox.ListImages("漫游L").Tag, ImgBox.ListImages("漫游R").Tag)
        ImgBox.ListImages("漫游L").Tag = ""
    End If

    If cMouseUsage("201").lngMouseKey = 1 Then
        ImgBox.ListImages("裁剪L").Tag = IIf(ImgBox.ListImages("裁剪L").Tag = "", ImgBox.ListImages("裁剪R").Tag, ImgBox.ListImages("裁剪L").Tag)
        ImgBox.ListImages("裁剪R").Tag = ""
        
        ImgBox.ListImages("框选L").Tag = IIf(ImgBox.ListImages("框选L").Tag = "", ImgBox.ListImages("框选R").Tag, ImgBox.ListImages("框选L").Tag)
        ImgBox.ListImages("框选R").Tag = ""
    Else
        ImgBox.ListImages("裁剪R").Tag = IIf(ImgBox.ListImages("裁剪R").Tag = "", ImgBox.ListImages("裁剪L").Tag, ImgBox.ListImages("裁剪R").Tag)
        ImgBox.ListImages("裁剪L").Tag = ""
        
        ImgBox.ListImages("框选R").Tag = IIf(ImgBox.ListImages("框选R").Tag = "", ImgBox.ListImages("框选L").Tag, ImgBox.ListImages("框选R").Tag)
        ImgBox.ListImages("框选L").Tag = ""
    End If


    If cMouseUsage("102").lngMouseKey = 1 Then
        ImgBox.ListImages("手动调窗L").Tag = IIf(ImgBox.ListImages("手动调窗L").Tag = "", ImgBox.ListImages("手动调窗R").Tag, ImgBox.ListImages("手动调窗L").Tag)
        ImgBox.ListImages("手动调窗R").Tag = ""
    Else
        ImgBox.ListImages("手动调窗R").Tag = IIf(ImgBox.ListImages("手动调窗R").Tag = "", ImgBox.ListImages("手动调窗L").Tag, ImgBox.ListImages("手动调窗R").Tag)
        ImgBox.ListImages("手动调窗L").Tag = ""
    End If

    If cMouseUsage("105").lngMouseKey = 1 Then
        ImgBox.ListImages("自适应调窗L").Tag = IIf(ImgBox.ListImages("自适应调窗L").Tag = "", ImgBox.ListImages("自适应调窗R").Tag, ImgBox.ListImages("自适应调窗L").Tag)
        ImgBox.ListImages("自适应调窗R").Tag = ""
    Else
        ImgBox.ListImages("自适应调窗R").Tag = IIf(ImgBox.ListImages("自适应调窗R").Tag = "", ImgBox.ListImages("自适应调窗L").Tag, ImgBox.ListImages("自适应调窗R").Tag)
        ImgBox.ListImages("自适应调窗L").Tag = ""
    End If
    
    If cMouseUsage("106").lngMouseKey = 1 Then
        ImgBox.ListImages("三维鼠标L").Tag = IIf(ImgBox.ListImages("三维鼠标L").Tag = "", ImgBox.ListImages("三维鼠标R").Tag, ImgBox.ListImages("三维鼠标L").Tag)
        ImgBox.ListImages("三维鼠标R").Tag = ""
    Else
        ImgBox.ListImages("三维鼠标R").Tag = IIf(ImgBox.ListImages("三维鼠标R").Tag = "", ImgBox.ListImages("三维鼠标L").Tag, ImgBox.ListImages("三维鼠标R").Tag)
        ImgBox.ListImages("三维鼠标L").Tag = ""
    End If

    If cMouseUsage("8").lngMouseKey = 1 Then
        ImgBox.ListImages("文字L").Tag = IIf(ImgBox.ListImages("文字L").Tag = "", ImgBox.ListImages("文字R").Tag, ImgBox.ListImages("文字L").Tag)
        ImgBox.ListImages("文字R").Tag = ""
    Else
        ImgBox.ListImages("文字R").Tag = IIf(ImgBox.ListImages("文字R").Tag = "", ImgBox.ListImages("文字L").Tag, ImgBox.ListImages("文字R").Tag)
        ImgBox.ListImages("文字L").Tag = ""
    End If

    If cMouseUsage("4").lngMouseKey = 1 Then
        ImgBox.ListImages("箭头L").Tag = IIf(ImgBox.ListImages("箭头L").Tag = "", ImgBox.ListImages("箭头R").Tag, ImgBox.ListImages("箭头L").Tag)
        ImgBox.ListImages("箭头R").Tag = ""
    Else
        ImgBox.ListImages("箭头R").Tag = IIf(ImgBox.ListImages("箭头R").Tag = "", ImgBox.ListImages("箭头L").Tag, ImgBox.ListImages("箭头R").Tag)
        ImgBox.ListImages("箭头L").Tag = ""
    End If

    If cMouseUsage("3").lngMouseKey = 1 Then
        ImgBox.ListImages("椭圆L").Tag = IIf(ImgBox.ListImages("椭圆L").Tag = "", ImgBox.ListImages("椭圆R").Tag, ImgBox.ListImages("椭圆L").Tag)
        ImgBox.ListImages("椭圆R").Tag = ""
    Else
        ImgBox.ListImages("椭圆R").Tag = IIf(ImgBox.ListImages("椭圆R").Tag = "", ImgBox.ListImages("椭圆L").Tag, ImgBox.ListImages("椭圆R").Tag)
        ImgBox.ListImages("椭圆L").Tag = ""
    End If

    If cMouseUsage("7").lngMouseKey = 1 Then
        ImgBox.ListImages("角度L").Tag = IIf(ImgBox.ListImages("角度L").Tag = "", ImgBox.ListImages("角度R").Tag, ImgBox.ListImages("角度L").Tag)
        ImgBox.ListImages("角度R").Tag = ""
    Else
        ImgBox.ListImages("角度R").Tag = IIf(ImgBox.ListImages("角度R").Tag = "", ImgBox.ListImages("角度L").Tag, ImgBox.ListImages("角度R").Tag)
        ImgBox.ListImages("角度L").Tag = ""
    End If
    
    If cMouseUsage("6").lngMouseKey = 1 Then
        ImgBox.ListImages("曲线L").Tag = IIf(ImgBox.ListImages("曲线L").Tag = "", ImgBox.ListImages("曲线R").Tag, ImgBox.ListImages("曲线L").Tag)
        ImgBox.ListImages("曲线R").Tag = ""
    Else
        ImgBox.ListImages("曲线R").Tag = IIf(ImgBox.ListImages("曲线R").Tag = "", ImgBox.ListImages("曲线L").Tag, ImgBox.ListImages("曲线R").Tag)
        ImgBox.ListImages("曲线L").Tag = ""
    End If
    
    If cMouseUsage("5").lngMouseKey = 1 Then
        ImgBox.ListImages("区域L").Tag = IIf(ImgBox.ListImages("区域L").Tag = "", ImgBox.ListImages("区域R").Tag, ImgBox.ListImages("区域L").Tag)
        ImgBox.ListImages("区域R").Tag = ""
    Else
        ImgBox.ListImages("区域R").Tag = IIf(ImgBox.ListImages("区域R").Tag = "", ImgBox.ListImages("区域L").Tag, ImgBox.ListImages("区域R").Tag)
        ImgBox.ListImages("区域L").Tag = ""
    End If
    
    If cMouseUsage("1").lngMouseKey = 1 Then
        ImgBox.ListImages("直线L").Tag = IIf(ImgBox.ListImages("直线L").Tag = "", ImgBox.ListImages("直线R").Tag, ImgBox.ListImages("直线L").Tag)
        ImgBox.ListImages("直线R").Tag = ""
        ImgBox.ListImages("血管测量L").Tag = IIf(ImgBox.ListImages("血管测量L").Tag = "", ImgBox.ListImages("血管测量R").Tag, ImgBox.ListImages("血管测量L").Tag)
        ImgBox.ListImages("血管测量R").Tag = ""
        ImgBox.ListImages("心胸比L").Tag = IIf(ImgBox.ListImages("心胸比L").Tag = "", ImgBox.ListImages("心胸比R").Tag, ImgBox.ListImages("心胸比L").Tag)
        ImgBox.ListImages("心胸比R").Tag = ""
    Else
        ImgBox.ListImages("直线R").Tag = IIf(ImgBox.ListImages("直线R").Tag = "", ImgBox.ListImages("直线L").Tag, ImgBox.ListImages("直线R").Tag)
        ImgBox.ListImages("直线L").Tag = ""
        ImgBox.ListImages("血管测量R").Tag = IIf(ImgBox.ListImages("血管测量R").Tag = "", ImgBox.ListImages("血管测量L").Tag, ImgBox.ListImages("血管测量R").Tag)
        ImgBox.ListImages("血管测量L").Tag = ""
        ImgBox.ListImages("心胸比R").Tag = IIf(ImgBox.ListImages("心胸比R").Tag = "", ImgBox.ListImages("心胸比L").Tag, ImgBox.ListImages("心胸比R").Tag)
        ImgBox.ListImages("心胸比L").Tag = ""
    End If

    If cMouseUsage("2").lngMouseKey = 1 Then
        ImgBox.ListImages("矩形L").Tag = IIf(ImgBox.ListImages("矩形L").Tag = "", ImgBox.ListImages("矩形R").Tag, ImgBox.ListImages("矩形L").Tag)
        ImgBox.ListImages("矩形R").Tag = ""
    Else
        ImgBox.ListImages("矩形R").Tag = IIf(ImgBox.ListImages("矩形R").Tag = "", ImgBox.ListImages("矩形L").Tag, ImgBox.ListImages("矩形R").Tag)
        ImgBox.ListImages("矩形L").Tag = ""
    End If

    If cMouseUsage("104").lngMouseKey = 1 Then
        ImgBox.ListImages("缩放L").Tag = IIf(ImgBox.ListImages("缩放L").Tag = "", ImgBox.ListImages("缩放R").Tag, ImgBox.ListImages("缩放L").Tag)
        ImgBox.ListImages("缩放R").Tag = ""
    Else
        ImgBox.ListImages("缩放R").Tag = IIf(ImgBox.ListImages("缩放R").Tag = "", ImgBox.ListImages("缩放L").Tag, ImgBox.ListImages("缩放R").Tag)
        ImgBox.ListImages("缩放L").Tag = ""
    End If
    
End Sub

Public Sub CreateMenu(ToolBars As Object, IconX As Integer, IconY As Integer)
    '------------------------------------------------
    '功能：                                  创建菜单
    '参数：
    '           IconX                        设置图标X大小
    '           IconY                        设置图标Y大小
    '返回：                                  无
    '上级函数或过程：                        frViewer_load
    '下级函数或过程：                        无
    '引用的外部参数：                        无
    '编制人：                                曾超 2005-6-27
    '------------------------------------------------
    Dim control As CommandBarControl
    Dim ControlFile As CommandBarPopup
    Dim ControlSelect As CommandBarPopup
    
    ToolBars.Options.UseDisabledIcons = True
    ToolBars.ActiveMenuBar.EnableDocking xtpFlagHideWrap
    '创建菜单
    '''''''''''''''''''''''''''''''''''''''文件菜单''''''''''''''''''''''''''''''''''''''''''''''
    Set ControlFile = ToolBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "文件(&F)", -1, False)
    With ControlFile.CommandBar.Controls
        .Add xtpControlButton, ID_File_Open, "打开(&O)"
        .Add xtpControlButton, ID_File_Close, "关闭序列(&C)"
        .Add xtpControlButton, ID_File_DelAllPhoto, "删除所有图像(&K)"
        .Add xtpControlButton, ID_File_DelReport, "删除报告图像(&D)"
        
        
        Set control = .Add(xtpControlButton, ID_File_SaveFile, "保存文件(&S)")
        control.BeginGroup = True
        
        .Add xtpControlButton, ID_File_SaveASFile, "另存文件(&A)", -1, False
        .Add xtpControlButton, ID_File_SaveToCD, "创建CD", -1, False
        .Add xtpControlButton, ID_File_SAveASReport, "保存报告图(&R)", -1, False
        
        Set ControlSelect = .Add(xtpControlPopup, ID_File_Send, "发送")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_File_Send_GetHost, "接收主机(&H)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_File_Send_OutPowerPoint, "输出到PowerPoint"
        ControlSelect.BeginGroup = True
        .Add xtpControlButton, ID_File_OpenDicomDir, "打开DICOMDIR"
        .Add xtpControlButton, ID_File_PhotoProperty, "图像属性(&I)"
        
        Set control = .Add(xtpControlButton, ID_File_Exit, "退出(&X)")
        control.BeginGroup = True
    End With
    ''''''''''''''''''''''''''''''''''''''''''视图菜单''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set ControlFile = ToolBars.ActiveMenuBar.Controls.Add(xtpControlPopup, ID_View, "视图(&V)", -1, False)
    With ControlFile.CommandBar.Controls
        .Add xtpControlButton, ID_View_UpSeries, "上一序列"
        .Add xtpControlButton, ID_View_DownSeries, "下一序列"
        .Add xtpControlButton, ID_View_Typeset, "版面安排(T)"
        Set control = .Add(xtpControlButton, ID_View_OneBrowse, "单序列观察")
        control.BeginGroup = True
        Set control = .Add(xtpControlButton, ID_View_PropertyShow, "属性显示(&P)")
        control.Checked = True
        Set control = .Add(xtpControlButton, ID_View_LableShow, "标注显示(&L)")
        control.Checked = True
        Set control = .Add(xtpControlButton, ID_View_ShowOverlay, "显示Overlay")
        control.Checked = True
        .Add xtpControlButton, ID_View_ShowMiniSeries, "显示序列缩略图(&M)"
        .Add xtpControlButton, ID_View_ViewAllSeries, "全序列观片"
        
        Set ControlSelect = .Add(xtpControlPopup, ID_View_PhotoSerial, "图像顺序(&S)")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_PhotoNumber, "图像号(&1)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_BedASC, "床位正序(&2)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_BedDESC, "床位逆序(&3)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_CollectionTime, "采集时间(&4)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_PhotoTime, "图像时间(&5)"
        
        Set ControlSelect = .Add(xtpControlPopup, ID_View_ShowScale, "显示比例(&Z)")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_AutoShow, "自适应(&O)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_50%, "50%"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_100%, "100%"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_150%, "150%"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_200%, "200%"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_250%, "250%"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_300%, "300%"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_400%, "400%"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_Custom, "自定义(&A)"
        ControlSelect.BeginGroup = True
        
        .Add xtpControlButton, ID_View_FullScreen, "全屏显示(&U)", -1, False
    End With
    ''''''''''''''''''''''''''''''''''''''''''''''动作菜单'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set ControlFile = ToolBars.ActiveMenuBar.Controls.Add(xtpControlPopup, ID_Active, "动作(&A)", -1, False)
    With ControlFile.CommandBar.Controls
        
        Set ControlSelect = .Add(xtpControlPopup, ID_Active_Select, "选择(&S)")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Select_OneSelect, "单幅选择(&O)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Select_SelectAllSerial, "选择所有序列(&S)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Acitve_Select_SelectAllPhoto, "图像全选(&A)"
        
        Set ControlSelect = .Add(xtpControlPopup, ID_Active_Also, "同步")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Also_Serial, "序列同步(&S)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Also_Photo, "图像同步(&I)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Also_ManualSerial, "手工序列同步"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Also_LockSerial, "锁定/解锁序列"
                    
        Set control = .Add(xtpControlButton, ID_Active_Shuttle, "穿梭(&T)")
        control.BeginGroup = True
        
        .Add xtpControlButton, ID_Active_Cruise, "漫游(&M)"
        .Add xtpControlButton, ID_Active_Cut, "裁剪"
        .Add xtpControlButton, ID_ACtive_FrameSelectImage, "框选图像"
        .Add xtpControlButton, ID_Active_Zoom, "缩放(&Z)"
        .Add xtpControlButton, ID_Active_ReSetAll, "恢复所有(&A)"
        .Add xtpControlButton, ID_ACtive_Mouse_Value, "在鼠标上显示CT值(&S)"
        .Add xtpControlButton, ID_Tool_NothinMouseState, "清除所有鼠标状态(ESC)"
        
        Set ControlSelect = .Add(xtpControlPopup, ID_Active_AdjustWindow, "调窗(&W)")
            ControlSelect.CommandBar.Controls.Add xtpControlSplitButtonPopup, ID_Active_AdjustWindow_HandAdjustWindow, "手动调窗"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_AdjustWindow_AutoAdjustWindow, "自适应调窗"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_AdjustWindow_HandAdjustWindow_Custom, "自定义(&A)"
            
        Set ControlSelect = .Add(xtpControlPopup, ID_Active_PointingLine, "定位线(&P)")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_PointingLine_ALL, "所有定位线(&O)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_PointingLine_FirstLast, "首尾定位线(&1)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_PointingLine_Now, "当前定位线(&2)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_PointingLine_3DLine, "3D鼠标定位(&M)"
        ControlSelect.BeginGroup = True
        
        Set ControlSelect = .Add(xtpControlPopup, ID_Active_Eddy, "旋转(&R)")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Eddy_LeftRight, "左右翻转(&X)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Eddy_TopButton, "垂直翻转(&Y)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Eddy_Left90, "左旋90°"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Eddy_Right90, "右旋90°"
        ControlSelect.BeginGroup = True
        
        .Add xtpControlButton, ID_Active_ReverseVideo, "反白"
        
        Set ControlSelect = .Add(xtpControlPopup, ID_Active_SieveLens, "滤镜(&A)")
            ControlSelect.CommandBar.Controls.Add xtpControlButtonPopup, ID_Active_SieveLens_Model, "滤镜模板"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_SieveLens_LancetMinus, "增强强度减少(&D)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_SieveLens_LancetAdd, "增强强度增加(&U)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_SieveLens_FlatnessMinus, "平滑减少"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_SieveLens_FlatnessAdd, "平滑增加"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Sievelens_LeftMoveMinus, "增强幅度减弱(&M)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Sievelens_LeftMoveAdd, "增强幅度增强(&T)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Sievelens_PhotoReset, "图像还原(&R)"
        
        
        Set ControlSelect = .Add(xtpControlPopup, ID_Active_Lable, "标注(&M)")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_Text, "文字(&T)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_Arrowhead, "箭头(&P)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_Ellipse, "椭圆(&E)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_Angle, "角度(&G)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_Curve, "曲线(&C)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_Area, "区域(&A)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_BeeLine, "直线(&B)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_Rect, "矩形(&R)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_VasMeasure, "血管狭窄测量"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_CadioThoracicRatio, "心胸比测量"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_AreaBeeLinePhoto, "区域直方图(&H)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_AdjustLine, "校准(&V)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_ClearLbale, "清除标注(&E)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_DelSelectLable, "删除标注(&D)"
        
    End With
    '''''''''''''''''''''''''''''''''''''''''''工具''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set ControlFile = ToolBars.ActiveMenuBar.Controls.Add(xtpControlPopup, ID_Tool, "工具(&T)", -1, False)
    
    With ControlFile.CommandBar.Controls
        .Add xtpControlButton, ID_Tool_Movie, "电影"
        .Add xtpControlButton, ID_Tool_Magnifier, "放大镜(&G)"
        
        Set control = .Add(xtpControlButton, ID_Tool_ArrowyCoronaryReset, "矢冠状重建(&V)")
        control.BeginGroup = True
        Set control = .Add(xtpControlButton, ID_Tool_SlopeReconstruction, "斜面重建(&S)")
        
        .Add xtpControlButton, ID_Tool_NumberMinusShadow, "数字减影(&D)"
        .Add xtpControlButton, ID_Tool_BogusColour, "伪彩观察(&C)"
        
        Set control = .Add(xtpControlButton, ID_Tool_FilmPrint, "胶片打印(&P)")
        Set ControlSelect = .Add(xtpControlPopup, ID_Tool_Film_AddSeries, "打印序列")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddSeries, "打印序列"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddImage, "打印图像"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddSelected, "打印所选图"
            Set ControlSelect = ControlSelect.CommandBar.Controls.Add(xtpControlButtonPopup, ID_Tool_Film_AddInterval, "间隔打印")
            ControlSelect.CommandBar.SetPopupToolBar True
            ControlSelect.CommandBar.Title = "间隔打印"
            ControlSelect.ToolTipText = "间隔打印当前序列"
        control.BeginGroup = True
        
        .Add xtpControlButton, ID_Tool_PhotoUnite, "图像拼接(&I)"
        '通过使用app.logmode来判断当前程序是在源程序的调试状态还是在exe文件的执行状态。
        'App.LogMode = 0为调试状态，如果是源程序的调试状态，则加入观片选项的菜单。在exe文件的执行状态，这个菜单不显示。
        If App.LogMode = 0 Then .Add xtpControlButton, ID_Tool_LableTool, "标注工具"
        
        Set control = .Add(xtpControlButton, ID_Tool_LookPhotoOption, "观片选项(&O)")
        control.BeginGroup = True
        
        Set ControlSelect = .Add(xtpControlPopup, ID_ToolBar, "工具栏(&B)")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_ToolBar_Left, "靠左(&L)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_ToolBar_Right, "靠右(&R)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_ToolBar_Top, "靠上(&T)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_ToolBar_Button, "靠下(&B)"
            
            Set control = ControlSelect.CommandBar.Controls.Add(xtpControlButton, ID_toolBar_16Icon, "16*16图标")
            control.BeginGroup = True
            
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_ToolBar_24Icon, "24*24图标"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_ToolBar_32Icon, "32*32图标"
            
            Set control = ControlSelect.CommandBar.Controls.Add(xtpControlButton, ID_ToolBar_Hide, "隐藏(&H)")
            control.BeginGroup = True
    End With
    ''''''''''''''''''''''''''''''''''''''''''帮助''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set ControlFile = ToolBars.ActiveMenuBar.Controls.Add(xtpControlPopup, ID_Help, "帮助(&H)", -1, False)
    With ControlFile.CommandBar.Controls
        .Add xtpControlButton, ID_Help_Help, "帮助(&H)"
        
        Set ControlSelect = .Add(xtpControlPopup, ID_Help_WebZLSOFT, "WEB上的中联")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Help_WebZLSOFT_WEB, "中联主页(&H)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Help_WebZLSOFT_Mail, "发送反馈(&K)"
        
        Set control = .Add(xtpControlButton, ID_Help_UpdateDB, "升级数据库(&U)")
        control.BeginGroup = True
        
        Set control = .Add(xtpControlButton, ID_Help_About, "关于(&A)")
        control.BeginGroup = True
    End With
    
    
    
    
    '增加快捷键
    'Ctrl
    ToolBars.KeyBindings.Add FCONTROL, Asc("O"), ID_File_Open
    ToolBars.KeyBindings.Add FCONTROL, Asc("Q"), ID_File_Close
    ToolBars.KeyBindings.Add FCONTROL, Asc("R"), ID_File_SAveASReport
    ToolBars.KeyBindings.Add FCONTROL, Asc("X"), ID_File_Exit
    ToolBars.KeyBindings.Add FCONTROL, Asc("J"), ID_View_Typeset
    ToolBars.KeyBindings.Add FCONTROL, Asc("L"), ID_Active_Also_LockSerial
    ToolBars.KeyBindings.Add FCONTROL, Asc("1"), ID_Active_Select_OneSelect
    ToolBars.KeyBindings.Add FCONTROL, Asc("M"), ID_View_ShowMiniSeries
    ToolBars.KeyBindings.Add FCONTROL, Asc("G"), ID_Tool_Magnifier
'    ToolBars.KeyBindings.Add FCONTROL, Asc("A"), ID_Active_Select_SelectAllSerial
    ToolBars.KeyBindings.Add FCONTROL, Asc("A"), ID_Active_Also_ManualSerial
    
    'Alt
    ToolBars.KeyBindings.Add FALT, Asc("T"), ID_Active_Shuttle
    ToolBars.KeyBindings.Add FALT, Asc("M"), ID_Active_Cruise
    ToolBars.KeyBindings.Add FALT, Asc("J"), ID_ACtive_FrameSelectImage
    ToolBars.KeyBindings.Add FALT, Asc("R"), ID_Active_ReSetAll
    ToolBars.KeyBindings.Add FALT, Asc("B"), ID_Active_ReverseVideo
    ToolBars.KeyBindings.Add FALT, Asc("H"), ID_ToolBar_Hide
    
    
    
    '创建工具条
    Dim ToolBar As CommandBar
    Dim ControlPopup As CommandBarPopup
    
    Set ToolBar = ToolBars.Add("主工具栏", xtpBarBottom)
    ToolBar.SetIconSize IconX, IconY
    
    With ToolBar.Controls
        .Add xtpControlButton, ID_File_SAveASReport, "保存报告图"
        .Add xtpControlButton, ID_File_Open, "打开"
        .Add xtpControlButton, ID_Tool_FilmPrint, "胶片输出"
        Set ControlSelect = .Add(xtpControlSplitButtonPopup, ID_Tool_Film_AddSeries, "打印序列")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddSeries, "打印序列"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddImage, "打印图像"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddSelected, "打印所选图"
            Set ControlSelect = ControlSelect.CommandBar.Controls.Add(xtpControlButtonPopup, ID_Tool_Film_AddInterval, "间隔打印")
            ControlSelect.CommandBar.SetPopupToolBar True
            ControlSelect.CommandBar.Title = "间隔打印"
            ControlSelect.ToolTipText = "间隔打印当前序列"
    End With
    
    
    Set ToolBar = ToolBars.Add("图像操作", xtpBarBottom)
    ToolBar.SetIconSize IconX, IconY
    With ToolBar.Controls
        .Add xtpControlButton, ID_Active_Eddy_LeftRight, "水平镜象"
        .Add xtpControlButton, ID_Active_Eddy_TopButton, "垂直镜象"
        .Add xtpControlButton, ID_Active_Eddy_Left90, "左转90度"
        .Add xtpControlButton, ID_Active_Eddy_Right90, "右转90度"
        .Add xtpControlButton, ID_Active_ReverseVideo, "反白"
        .Add xtpControlButton, ID_Tool_NumberMinusShadow, "DSA数字减影"
    End With
    
    Set ToolBar = ToolBars.Add("测量工具栏", xtpBarBottom)
    ToolBar.SetIconSize IconX, IconY
    With ToolBar.Controls
        .Add xtpControlButton, ID_Tool_NothinMouseState, "鼠标"
        .Add xtpControlButton, ID_ACtive_Mouse_Value, "在鼠标上显示CT值"
        .Add xtpControlButton, ID_Active_Lable_Text, "文字"
        .Add xtpControlButton, ID_Active_Lable_Arrowhead, "箭头"
        .Add xtpControlButton, ID_Active_Lable_Ellipse, "椭圆"
        .Add xtpControlButton, ID_Active_Lable_Angle, "角度"
        .Add xtpControlButton, ID_Active_Lable_Curve, "曲线"
        .Add xtpControlButton, ID_Active_Lable_Area, "区域"
        .Add xtpControlButton, ID_Active_Lable_BeeLine, "直线"
        .Add xtpControlButton, ID_Active_Lable_Rect, "矩形"
        .Add xtpControlButton, ID_Active_Lable_VasMeasure, "血管狭窄测量"
        .Add xtpControlButton, ID_Active_Lable_CadioThoracicRatio, "心胸比测量"
        .Add xtpControlButton, ID_Active_Lable_ClearLbale, "清除标注"
        .Add xtpControlButton, ID_Active_Lable_AdjustLine, "校准"
    End With

    Set ToolBar = ToolBars.Add("多平面工具栏", xtpBarBottom)
    ToolBar.SetIconSize IconX, IconY
    With ToolBar.Controls
        .Add xtpControlButton, ID_Active_PointingLine_ALL, "显示所有定位线"
        .Add xtpControlButton, ID_Active_PointingLine_FirstLast, "显示首尾定位线"
        .Add xtpControlButton, ID_Active_PointingLine_Now, "显示当前定位线"
        .Add xtpControlButton, ID_Active_PointingLine_3DLine, "三维鼠标"
        .Add xtpControlButton, ID_Tool_ArrowyCoronaryReset, "矢/冠状位重建"
        .Add xtpControlButton, ID_Tool_SlopeReconstruction, "斜面重建"
    End With

    Set ToolBar = ToolBars.Add("对象分析", xtpBarBottom)
    ToolBar.SetIconSize IconX, IconY
    With ToolBar.Controls
        .Add xtpControlButton, ID_ACtive_FrameSelectImage, "框选图像"
        .Add xtpControlButton, ID_Active_Also_Photo, "图像格式同步"
        .Add xtpControlButton, ID_Active_Also_Serial, "序列间图像位置同步"
        .Add xtpControlButton, ID_Active_Also_ManualSerial, "手工序列同步"
        .Add xtpControlButton, ID_Active_Also_LockSerial, "锁定/解锁序列"
        .Add xtpControlButton, ID_View_ShowMiniSeries, "显示序列缩略图"
        .Add xtpControlButton, ID_View_ViewAllSeries, "全序列观片"
    End With

    Set ToolBar = ToolBars.Add("通用工具栏", xtpBarBottom)
    ToolBar.Closeable = False
    ToolBar.SetIconSize IconX, IconY
    With ToolBar.Controls
        .Add xtpControlButton, ID_Tool_Magnifier, "放大镜"
        Set ControlPopup = .Add(xtpControlSplitButtonPopup, ID_Active_AdjustWindow_HandAdjustWindow, "手控调窗")

        .Add xtpControlButton, ID_Active_Cruise, "漫游"

        Set ControlPopup = .Add(xtpControlSplitButtonPopup, ID_Active_Zoom, "缩放")
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_AutoShow, "自适应"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_50%, "50%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_100%, "100%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_150%, "150%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_200%, "200%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_250%, "250%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_300%, "300%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_400%, "400%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_Custom, "自定义(&A)"

        Set ControlPopup = .Add(xtpControlSplitButtonPopup, ID_Active_Shuttle, "穿梭")
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_PhotoNumber, "图像号"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_BedASC, "床位正序"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_BedDESC, "床位逆序"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_CollectionTime, "采集时间"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_PhotoTime, "图像时间"

        .Add xtpControlButton, ID_Tool_Movie, "电影播放"
        .Add xtpControlButton, ID_Active_Select_SelectAllSerial, "选择所有序列"
        .Add xtpControlButton, ID_Acitve_Select_SelectAllPhoto, "选择序列中所有的图像"
        .Add xtpControlButton, ID_View_UpSeries, "上一序列"
        .Add xtpControlButton, ID_View_DownSeries, "下一序列"
        .Add xtpControlButton, ID_View_Typeset, "版面设计"
        .Add xtpControlButton, ID_View_FullScreen, "全屏显示"
        Set control = .Add(xtpControlButton, ID_View_PropertyShow, "显/隐病人信息")
        control.Checked = True
        .Add xtpControlButton, ID_Active_ReSetAll, "全部恢复"
        .Add xtpControlButton, ID_View_OneBrowse, "浏览/观察模式"
    End With
    
    Set ToolBar = ToolBars.Add("图像增强", xtpBarBottom)
    ToolBar.SetIconSize IconX, IconY
    With ToolBar.Controls
        Set ControlPopup = .Add(xtpControlSplitButtonPopup, ID_Active_SieveLens_Model, "滤镜模板")
        .Add xtpControlButton, ID_Active_SieveLens_LancetMinus, "边缘增强强度减少"
        .Add xtpControlButton, ID_Active_SieveLens_LancetAdd, "边缘增强强度增加"
        .Add xtpControlButton, ID_Active_Sievelens_LeftMoveMinus, "边缘增强幅度减少"
        .Add xtpControlButton, ID_Active_Sievelens_LeftMoveAdd, "边缘增强幅度增加"
        .Add xtpControlButton, ID_Active_SieveLens_FlatnessMinus, "平滑减少"
        .Add xtpControlButton, ID_Active_SieveLens_FlatnessAdd, "平滑增加"
        .Add xtpControlButton, ID_Active_Sievelens_PhotoReset, "图像复原"
        .Add xtpControlButton, ID_Tool_BogusColour, "伪彩"
    End With
    
    ToolBars.EnableCustomization True
    
End Sub

Public Function StatusBarTip(control As CommandBarControl) As String
'------------------------------------------------
'功能：在状态栏显示菜单的简单帮助
'参数：Control--显示帮助的菜单控件
'返回：帮助信息
'上级函数或过程：frmViewer.ComToolBar_ControlSelected
'下级函数或过程：无
'引用的外部参数：无
'编制人：曾超
'------------------------------------------------
    If control Is Nothing Then
        StatusBarTip = ""
        Exit Function
    End If
    Select Case control.Id
        ''''''''''''''''''''''''''文件菜单'''''''''''''''''''''''''''''''''''
        Case ID_File_Open                                                               '打开文件
            StatusBarTip = "打开新的图象文件进行观察"
            
        Case ID_File_Close                                                              '关闭序列
            StatusBarTip = "关闭当前序列图像，关闭后可以通过排版再次调出序列图像"
            
        Case ID_File_DelAllPhoto                                                        '删除所有图像
            StatusBarTip = "删除所有序列中图像"
            
        Case ID_File_DelReport                                                          '删除报告图像
            StatusBarTip = "删除报告图像"
            
        Case ID_File_SaveFile                                                           '保存文件
            StatusBarTip = "保存文件"
            
        Case ID_File_SaveASFile                                                         '另存文件
            StatusBarTip = "将当前选中的图象另存为文件"
            
        Case ID_File_SaveToCD                                                           '创建CD
            StatusBarTip = "将当前选中的图象保存到CD缓冲区"
            
        Case ID_File_SAveASReport                                                       '保存报告图像
            StatusBarTip = "将当前选中的图象保存为报告图像"
            
        Case ID_File_Send_GetHost                                                       '接收主机
            StatusBarTip = "发送到接收主机"
            
        Case ID_File_Send_OutPowerPoint                                                 '输出到PowerPoint
            StatusBarTip = "输出到PowerPoint"
            
        Case ID_File_OpenDicomDir                                                       '打开DICOMDIR
            StatusBarTip = "打开DICOMDIR中的图像"
            
        Case ID_File_PhotoProperty                                                      '图像属性
            StatusBarTip = "查看包含在图象中的病人、检查、序列和图象的相关信息"
            
        Case ID_File_Exit                                                               '退出
            StatusBarTip = "退出"
            
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''视图''''''''''''''''''''''''''''''''''''''''
        Case ID_View_UpSeries                                                           '上一序列
            StatusBarTip = "切换到上一个序列"
            
        Case ID_View_DownSeries                                                         '下一序列
            StatusBarTip = "切换到下一个序列"
            
        Case ID_View_Typeset                                                            '版面安排
            StatusBarTip = "调整打开序列及序列内图象的屏幕显示排列布局"
            
        Case ID_View_OneBrowse                                                          '浏览观察模式
            StatusBarTip = "使用浏览模式或是观察模式察看图像"
            
        Case ID_View_PropertyShow                                                       '图像上病人信息显示
            StatusBarTip = "显示或隐藏病人基本信息"
            
        Case ID_View_LableShow                                                          '标注显示
            StatusBarTip = "显示/隐藏标注信息"
            
        Case ID_View_ShowMiniSeries                                                     '显示序列缩略图
            StatusBarTip = "显示/隐藏序列缩略图"
            
        Case ID_View_PhotoSerial_PhotoNumber                                            '图像顺序_图像号
            StatusBarTip = "以图像号顺序浏览"
            
        Case ID_View_PhotoSerial_BedASC                                                 '床位正序
            StatusBarTip = "以床位正序顺序浏览"
            
        Case ID_View_PhotoSerial_BedDESC                                                '床位逆序
            StatusBarTip = "以床位逆序顺序浏览"
            
        Case ID_View_PhotoSerial_CollectionTime                                         '采集时间
            StatusBarTip = "以采集时间顺序浏览"
            
        Case ID_View_PhotoSerial_PhotoTime                                              '图像时间
            StatusBarTip = "以图像时间顺序浏览"
            
        Case ID_View_ShowScale_AutoShow                                                 '自适应
            StatusBarTip = "以屏幕合适的大小显示图像"
            
        Case ID_View_ShowScale_50%                                                      '50%
            StatusBarTip = "以50%大小图像显示"
            
        Case ID_View_ShowScale_100%                                                     '100%
            StatusBarTip = "以图像正常大小显示"
            
        Case ID_View_showScale_150%                                                     '150%
            StatusBarTip = "以150%图像大小显示"
            
        Case ID_View_ShowScale_200%                                                     '200%
            StatusBarTip = "以200%大小图像显示"
            
        Case ID_View_showScale_250%                                                     '250%
            StatusBarTip = "以250%大小图像显示"
            
        Case ID_View_showScale_300%                                                     '300%
            StatusBarTip = "以300%大小图像显示"
            
        Case ID_View_showScale_400%                                                     '400%
            StatusBarTip = "以400%大小图像显示"
            
        Case ID_View_ShowScale_Custom                                                   '自定义
            StatusBarTip = "自定义图像显示大小"
            
        Case ID_View_FullScreen                                                         '全屏显示
            StatusBarTip = "全屏幕显示图象进行观察"
            
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''动作''''''''''''''''''''''''''''''''''''
        Case ID_Active_Select_OneSelect                                                 '单幅选择
            StatusBarTip = "选择当前图像"
        
        Case ID_Active_Select_SelectAllSerial                                           '选择所有序列
            StatusBarTip = "选中打开的所有序列，以便为其他操作作准备"
            
        Case ID_Acitve_Select_SelectAllPhoto                                            '选择所有图像
            StatusBarTip = "选择或取消对当前序列的所有图象的选择标志"
            
        Case ID_Active_Also_Serial                                                      '序列同步
            StatusBarTip = "对当前打开的序列做图像位置的同步操作"
            
        Case ID_Active_Also_ManualSerial                                                '手工序列同步
            StatusBarTip = "手工对当前打开的序列做图像位置的同步操作"
        
        Case ID_Active_Also_LockSerial                                                  '锁定序列
            StatusBarTip = "锁定或解锁序列，设置该序列是否参加手动序列同步，与“键盘Ctrl+鼠标左键单击”功能相同"
            
        Case ID_Active_Also_Photo                                                       '图像同步
            StatusBarTip = "对当前序列内的图像操作结果做同步"
            
        Case ID_Active_Shuttle                                                          '穿梭
            StatusBarTip = "在观察区内通过鼠标的移动直接切换图像" & funcGetShiftStr(cMouseUsage("101").lngShift)
            
        Case ID_Active_Cruise                                                           '漫游
            StatusBarTip = "在观察区内移动图象的位置，以便于更好地观察" & funcGetShiftStr(cMouseUsage("103").lngShift)

        Case ID_Active_Cut                                                              '裁剪
            StatusBarTip = "对图像进行裁剪" & funcGetShiftStr(cMouseUsage("201").lngShift)
            
        Case ID_ACtive_FrameSelectImage                                                 '框选
            StatusBarTip = "拖动矩形框选择局部图像" & funcGetShiftStr(cMouseUsage("201").lngShift)
                        
        Case ID_Active_Zoom                                                             '缩放
            StatusBarTip = "在观察区内缩小或放大图像" & funcGetShiftStr(cMouseUsage("104").lngShift)
            
        Case ID_Active_ReSetAll                                                         '恢复所有
            StatusBarTip = "取消调窗、漫游等操作，恢复图象原始状态"
            
        Case ID_Active_AdjustWindow_HandAdjustWindow                                    '手动调窗
            StatusBarTip = "进入手动控制的图象窗宽窗位调节模式" & funcGetShiftStr(cMouseUsage("102").lngShift)
            
        Case ID_Active_AdjustWindow_AutoAdjustWindow                                    '自适应调窗
            StatusBarTip = "进入自适调窗模式，通过选择一个区域，进行自适应调整" & funcGetShiftStr(cMouseUsage("105").lngShift)
            
        Case ID_Active_AdjustWidnow_CustomAdjustWindow                                  '自定义调窗
            StatusBarTip = "输入合适的窗宽窗位进行调节"
            
        Case ID_Active_PointingLine_ALL                                                 '所有定位线
            StatusBarTip = "显示序列图象的所有定位线"
            
        Case ID_Active_PointingLine_FirstLast                                           '首位定位线
            StatusBarTip = "显示序列图象的首尾定位线"
            
        Case ID_Active_PointingLine_Now                                                 '当前定位线
            StatusBarTip = "显示当前图象对应的定位线"
            
        Case ID_Active_PointingLine_3DLine                                              '3D鼠标
            StatusBarTip = "显示当前图象鼠标指向的三维对应位置点" & funcGetShiftStr(cMouseUsage("106").lngShift)
            
        Case ID_Active_Eddy_LeftRight                                                   '左右旋转
            StatusBarTip = "对图象进行左右翻转后进行观察"
            
        Case ID_Active_Eddy_TopButton                                                   '垂直旋转
            StatusBarTip = "对图象进行垂直翻转后进行观察"
            
        Case ID_Active_Eddy_Left90                                                      '左旋90
            StatusBarTip = "对图象进行左旋90°后进行观察"
            
        Case ID_Active_Eddy_Right90                                                     '右旋90
            StatusBarTip = "对图象进行右旋90°后进行观察"
            
        Case ID_Active_ReverseVideo                                                     '反白
            StatusBarTip = "对当前图象及其同步的其他图象进行黑白反转观察"
        
        Case ID_Active_SieveLens_Model                                                  '滤镜模板
            StatusBarTip = "应用预先设置好的滤镜模板"
                 
        Case ID_Active_SieveLens_LancetMinus                                            '锐化减少
            StatusBarTip = "降低图象增强强度"
            
        Case ID_Active_SieveLens_LancetAdd                                              '锐化增加
            StatusBarTip = "增加图象增强强度"
            
        Case ID_Active_SieveLens_FlatnessMinus                                          '平滑减少
            StatusBarTip = "降低图像平滑效果"
            
        Case ID_Active_SieveLens_FlatnessAdd                                            '平滑增加
            StatusBarTip = "增加图像平滑效果"
            
        Case ID_Active_Sievelens_LeftMoveMinus                                          '左移增弱
            StatusBarTip = "降低图象增强幅度"
            
        Case ID_Active_Sievelens_LeftMoveAdd                                            '左移增加
            StatusBarTip = "增加图象增强幅度"
            
        Case ID_Active_Sievelens_PhotoReset                                             '图像还原
            StatusBarTip = "取消滤镜增强观察效果，恢复图象原始状态"
            
        Case ID_Active_Lable_Text                                                       '文字
            StatusBarTip = "添加“文字”类型的标注" & funcGetShiftStr(cMouseUsage("8").lngShift)
            
        Case ID_Active_Lable_Arrowhead                                                  '箭头
            StatusBarTip = "添加“箭头”类型的标注" & funcGetShiftStr(cMouseUsage("4").lngShift)
         
        Case ID_Active_Lable_Ellipse                                                    '椭圆
            StatusBarTip = "添加“椭圆”类型的标注" & funcGetShiftStr(cMouseUsage("3").lngShift)
        
        Case ID_Active_Lable_Angle                                                      '角度
            StatusBarTip = "添加“角度”类型的标注" & funcGetShiftStr(cMouseUsage("7").lngShift)
        
        Case ID_Active_Lable_Curve                                                      '曲线
            StatusBarTip = "添加“曲线”类型的标注" & funcGetShiftStr(cMouseUsage("6").lngShift)
        
        Case ID_Active_Lable_Area                                                       '区域
            StatusBarTip = "添加任意的封闭区域形式的标注" & funcGetShiftStr(cMouseUsage("5").lngShift)
        
        Case ID_Active_Lable_BeeLine                                                    '直线
            StatusBarTip = "添加“直线”类型的标注" & funcGetShiftStr(cMouseUsage("1").lngShift)
        
        Case ID_Active_Lable_Rect                                                       '矩形
            StatusBarTip = "添加“矩形”类型的标注" & funcGetShiftStr(cMouseUsage("2").lngShift)
        
        Case ID_Active_Lable_AreaBeeLinePhoto                                           '区域直方图
            StatusBarTip = "对选中的矩形椭圆等区域的灰度情况直方图对比"
            
        Case ID_Active_Lable_VasMeasure                                                 '血管狭窄测量
            StatusBarTip = "对图像进行血管狭窄测量" & funcGetShiftStr(cMouseUsage("1").lngShift)
            
        Case ID_Active_Lable_CadioThoracicRatio                                         '心胸比测量
            StatusBarTip = "对心脏和胸廓比例进行测量" & funcGetShiftStr(cMouseUsage("1").lngShift)
            
        Case ID_Active_Lable_AdjustLine                                                 '校准
            StatusBarTip = "对选中的直线进行长度的手工校准，修改标注"
            
        Case ID_Active_Lable_ClearLbale                                                 '清除所有标注
            StatusBarTip = "清除所有的标注"
            
        Case ID_Active_Lable_DelSelectLable                                             '删除当前标注
            StatusBarTip = "删除当前选中的标注与测量"
            
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''工具菜单''''''''''''''''''''''''''''''''''''
        Case ID_Tool_Movie                                                              '电影
            StatusBarTip = "按循环或钟摆方式连续播放多帧或多幅图象"
            
        Case ID_Tool_Magnifier                                                          '放大镜
            StatusBarTip = "使用放大镜，对图象进行局部的放大缩小观察"
            
        Case ID_Tool_ArrowyCoronaryReset                                                '矢冠状重建
            StatusBarTip = "对图象进行矢冠状位二维重建观察"
            
        Case ID_Tool_SlopeReconstruction                                                '斜面重建
            StatusBarTip = "对图象进行斜面二维重建观察"
            
        Case ID_Tool_NumberMinusShadow                                                  '数字减影
            StatusBarTip = "对图象进行数字减影的观察"
            
        Case ID_Tool_BogusColour                                                        '伪彩
            StatusBarTip = "设置并以伪彩色方式观察图象"
            
        Case ID_Tool_FilmPrint                                                          '胶片打印
            StatusBarTip = "进入对选中序列和图象的胶片打印处理"
            
        Case ID_Tool_PhotoUnite                                                         '图像拼接
            StatusBarTip = "几个同类型间的图像拼接"
            
        Case ID_Tool_LableTool                                                          '标注工具
            StatusBarTip = "标注工具"
            
        Case ID_Tool_LookPhotoOption                                                    '观片选项
            StatusBarTip = "观片工作站的基础设置"
            
        Case ID_ToolBar_Left                                                            '工具条靠左
            StatusBarTip = "工具条靠左摆放"
            
        Case ID_ToolBar_Right                                                           '工具条靠右
            StatusBarTip = "工具条靠右摆放"
        
        Case ID_ToolBar_Top                                                             '工具条靠上
            StatusBarTip = "工具条靠上摆放"
            
        Case ID_ToolBar_Button                                                          '工具条靠下
            StatusBarTip = "工具条靠下摆放"
            
        Case ID_toolBar_16Icon                                                          '工具条图标16*16显示
            StatusBarTip = "工具条以16*16图标显示"
            
        Case ID_ToolBar_24Icon                                                          '工具条图标24*24显示
            StatusBarTip = "工具条以24*24图标显示"
        
        Case ID_ToolBar_32Icon                                                          '工具条图标32*32显示
            StatusBarTip = "工具条以32*32图标显示"
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''帮助'''''''''''''''''''''''''''''''''''''
        Case ID_Help_Help                                                               '帮助
            StatusBarTip = "观片站帮助"
        
        Case ID_Help_WebZLSOFT_WEB                                                      '中联主页
            StatusBarTip = "打开中联主页"
        
        Case ID_Help_WebZLSOFT_Mail                                                     '发送反馈
            StatusBarTip = "发送反馈"
            
        Case ID_Help_About                                                              '关于
            StatusBarTip = "关于"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End Select
End Function

Private Function funcGetShiftStr(lngShift As Long) As String
'------------------------------------------------
'功能：通过给定shift状态数值，解析出用字符串表示的shift状态。
'参数：lngShift--表示shifit状态的数值
'返回：用字符串表示的shift状态。
'上级函数或过程：mdlTools.StatusBarTip
'下级函数或过程：无
'引用的外部参数：无
'编制人：黄捷
'------------------------------------------------
    'shift 的用法，shift,ctrl,alt 分别用1，2，4表示，通过累加实现
    funcGetShiftStr = ""
    If lngShift - 4 >= 0 Then
        funcGetShiftStr = " Alt "
        lngShift = lngShift - 4
    End If
    If lngShift - 2 >= 0 Then
        funcGetShiftStr = funcGetShiftStr & " Ctrl "
        lngShift = lngShift - 2
    End If
    If lngShift = 1 Then
        funcGetShiftStr = funcGetShiftStr & " Shift "
    End If
End Function

Public Sub ArrayToolBar(ToolBars As Object, frmTop As Long, frmLeft As Long, frmHeight As Long, frmWidth As Long)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能               重新按一定的顺序摆放工具栏位置
    '参数
    '    ToolBars       工具条控件
    '    frmTop         当前窗体Top
    '    frmLeft        当前窗体Left
    '    frmWidth       当前窗体Width
    '    frmHieht       当前窗体Height
    '返回               无
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim MenuToolBar As CommandBar                           '菜单
    Dim MainToolBar As CommandBar                           '主工具条
    Dim PhotoToolBar As CommandBar                          '图像处理工具条
    Dim ScaleToolBar As CommandBar                          '测量工具条
    Dim PlaneToolBar As CommandBar                          '平面工具条
    Dim ObjectToolBar As CommandBar                         '对象工具条
    Dim CommToolBar As CommandBar                           '通用工具条
    Dim PhotoStrongToolBar As CommandBar                    '图像增强工具条
    Dim NowPosiTion As Integer                              '当前工具条位置
    Const intState As Integer = 360                         '状态栏高
    Dim ToolBarLeft As Long                                 '工具栏最左边的位置
    Dim ToolBarTop As Long                                  '工具栏最上边的位置
    
    Dim OldTop As Long, OldLeft As Long, OldRight As Long, OldBottom As Long                    '上一工具条位置
    Dim NowTop As Long, NowLeft As Long, NowRight As Long, NowBottom As Long                    '当前工具条位置
    
    
    
    blfrmRefresh = False
    
    '主工具条
    Set MainToolBar = ToolBars.Item(2)
    NowPosiTion = MainToolBar.Position
    Set MenuToolBar = ToolBars.Item(1)
    MenuToolBar.GetWindowRect NowLeft, NowTop, NowRight, NowBottom
    ToolBarLeft = NowLeft
    ToolBarTop = NowTop
    Select Case MainToolBar.Position
        Case 0
            ToolBars.DockToolBar MainToolBar, ToolBarLeft, NowBottom, NowPosiTion
        Case 1
            ToolBars.DockToolBar MainToolBar, ToolBarLeft, frmHeight + frmTop, NowPosiTion
        Case 2
            ToolBars.DockToolBar MainToolBar, ToolBarLeft, ToolBarTop, NowPosiTion
        Case 3
            ToolBars.DockToolBar MainToolBar, frmLeft + frmWidth, ToolBarTop, NowPosiTion
    End Select
    
    ToolBars.RecalcLayout
    
    '对象分析工具条
    MainToolBar.GetWindowRect OldLeft, OldTop, OldRight, OldBottom
    
    
    Set ObjectToolBar = ToolBars.Item(6)
    ObjectToolBar.GetWindowRect NowLeft, NowTop, NowRight, NowBottom
    Select Case ObjectToolBar.Position
        Case 0
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar ObjectToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar ObjectToolBar, ToolBarLeft, OldBottom, NowPosiTion
            End If
        Case 1
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar ObjectToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar ObjectToolBar, ToolBarLeft, OldTop, NowPosiTion
            End If
        Case 2
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar ObjectToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar ObjectToolBar, OldRight, ToolBarTop, NowPosiTion
            End If
        Case 3
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar ObjectToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar ObjectToolBar, OldLeft - (OldRight - OldLeft), ToolBarTop, NowPosiTion
            End If
    End Select
    ToolBars.RecalcLayout
    
    '图像处理工具条
    ObjectToolBar.GetWindowRect OldLeft, OldTop, OldRight, OldBottom
    
    
    
    Set PhotoToolBar = ToolBars.Item(3)
    PhotoToolBar.GetWindowRect NowLeft, NowTop, NowRight, NowBottom
    Select Case PhotoToolBar.Position
        Case 0
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar PhotoToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar PhotoToolBar, ToolBarLeft, OldBottom, NowPosiTion
            End If
        Case 1
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar PhotoToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar PhotoToolBar, ToolBarLeft, OldTop, NowPosiTion
            End If
        Case 2
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar PhotoToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar PhotoToolBar, OldRight, ToolBarTop, NowPosiTion
            End If
        Case 3
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar PhotoToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar PhotoToolBar, OldLeft - (OldRight - OldLeft), ToolBarTop, NowPosiTion
            End If
    End Select
    ToolBars.RecalcLayout

    '测量工具条
    PhotoToolBar.GetWindowRect OldLeft, OldTop, OldRight, OldBottom
    
    
    
    Set ScaleToolBar = ToolBars.Item(4)
    ScaleToolBar.GetWindowRect NowLeft, NowTop, NowRight, NowBottom
    Select Case ScaleToolBar.Position
        Case 0
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar ScaleToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar ScaleToolBar, ToolBarLeft, OldBottom, NowPosiTion
            End If
        Case 1
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar ScaleToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar ScaleToolBar, ToolBarLeft, OldTop, NowPosiTion
            End If
        Case 2
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar ScaleToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar ScaleToolBar, OldRight, ToolBarTop, NowPosiTion
            End If
        Case 3
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar ScaleToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar ScaleToolBar, OldLeft - (OldRight - OldLeft), ToolBarTop, NowPosiTion
            End If
    End Select
    ToolBars.RecalcLayout

    '平面工具条
    ScaleToolBar.GetWindowRect OldLeft, OldTop, OldRight, OldBottom
    
    
    Set PlaneToolBar = ToolBars.Item(5)
    PlaneToolBar.GetWindowRect NowLeft, NowTop, NowRight, NowBottom
    Select Case ScaleToolBar.Position
        Case 0
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar PlaneToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar PlaneToolBar, ToolBarLeft, OldBottom, NowPosiTion
            End If
        Case 1
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar PlaneToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar PlaneToolBar, ToolBarLeft, OldTop, NowPosiTion
            End If
        Case 2
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar PlaneToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar PlaneToolBar, OldRight, ToolBarTop, NowPosiTion
            End If
        Case 3
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar PlaneToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar PlaneToolBar, OldLeft - (OldRight - OldLeft), ToolBarTop, NowPosiTion
            End If
    End Select
    ToolBars.RecalcLayout

    '通用工具条
    PlaneToolBar.GetWindowRect OldLeft, OldTop, OldRight, OldBottom
    
    
    
    Set CommToolBar = ToolBars.Item(7)
    CommToolBar.GetWindowRect NowLeft, NowTop, NowRight, NowBottom
    Select Case ScaleToolBar.Position
        Case 0
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar CommToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar CommToolBar, ToolBarLeft, OldBottom, NowPosiTion
            End If
        Case 1
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar CommToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar CommToolBar, ToolBarLeft, OldTop, NowPosiTion
            End If
        Case 2
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar CommToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar CommToolBar, OldRight, ToolBarTop, NowPosiTion
            End If
        Case 3
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar CommToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar CommToolBar, OldLeft - (OldRight - OldLeft), ToolBarTop, NowPosiTion
            End If
    End Select
    
    ToolBars.RecalcLayout
    
    '通用工具条
    CommToolBar.GetWindowRect OldLeft, OldTop, OldRight, OldBottom
    
    
    
    Set PhotoStrongToolBar = ToolBars.Item(8)
    PhotoStrongToolBar.GetWindowRect NowLeft, NowTop, NowRight, NowBottom
    Select Case ScaleToolBar.Position
        Case 0
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar PhotoStrongToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar PhotoStrongToolBar, ToolBarLeft, OldBottom, NowPosiTion
            End If
        Case 1
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar PhotoStrongToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar PhotoStrongToolBar, ToolBarLeft, OldTop, NowPosiTion
            End If
        Case 2
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar PhotoStrongToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar PhotoStrongToolBar, OldRight, ToolBarTop, NowPosiTion
            End If
        Case 3
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar PhotoStrongToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar PhotoStrongToolBar, OldLeft - (OldRight - OldLeft), ToolBarTop, NowPosiTion
            End If
    End Select
    
    blfrmRefresh = True
    ToolBars.RecalcLayout
End Sub


Sub ReplaceToolBarIcon(ObjToolBar As Object, imgList As ImageList, IconX As Integer, IconY As Integer)
    '------------------------------------------------
    '功能：                                  替换当前图标为16,24,32
    '参数：
    '           objToolbar                   工具条对象
    '           imglist                      图标更表对象
    '           IconX                        设置图标X大小
    '           IconY                        设置图标Y大小
    '返回：                                  无
    '上级函数或过程：                        ComToolBar_Execute
    '下级函数或过程：                        无
    '引用的外部参数：                        无
    '编制人：                                曾超 2005-6-29
    '------------------------------------------------
    Dim i As Integer
    For i = 2 To ObjToolBar.Count
        ObjToolBar.Item(i).SetIconSize IconX, IconY
    Next
    ObjToolBar.AddImageList imgList
End Sub

Sub PutToolbar(ObjToolBar As Object, Position As Integer)
    '------------------------------------------------
    '功能：                                  摆放工具条到Top,Left,Right,Bottom
    '参数：
    '           objToolbar                   工具条对象
    '           position                     摆放位置 0=Top,1=Bottom 2=Left 3=Right
    '返回：                                  无
    '上级函数或过程：                        ComToolBar_Execute
    '下级函数或过程：                        无
    '引用的外部参数：                        无
    '编制人：                                曾超 2005-6-29
    '------------------------------------------------
    Dim i As Integer
    For i = 2 To ObjToolBar.Count
        ObjToolBar.Item(i).Position = Position
    Next
End Sub

Public Sub WriteLog(ByVal ErrorType As Integer, ErrorNum As Long, ErrorDesc As String)
    Dim strSQL As String
    
    On Error GoTo errh
    
    If blLocalRun = False Then Exit Sub
    If cnAccess.State = adStateClosed Then Exit Sub
    
    strSQL = "Insert Into 错误日志(产生时间,错误类型,错误号,错误信息) " & _
        "Values(cDate('" & Date & " " & Time() & "')," & ErrorType & "," & ErrorNum & ",'" & Replace(ErrorDesc, "'", "''") & "')"
    cnAccess.Execute strSQL
    
    Exit Sub
errh:
    MsgBox "发生错误:" & err.Description, vbExclamation, gstrSysName
End Sub


