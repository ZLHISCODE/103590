VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmInDoctorView 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "住院一览"
   ClientHeight    =   8625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13440
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7065
      Left            =   0
      ScaleHeight     =   7065
      ScaleWidth      =   12000
      TabIndex        =   0
      Top             =   480
      Width           =   12000
   End
   Begin XtremeCommandBars.CommandBars cbsSub 
      Left            =   120
      Top             =   840
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmInDoctorView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'事件
Public Event ViewPACSImage(ByVal 医嘱ID As Long) '要求进行观片
Public Event ResizeForm(ByVal bytFunc As Long)  '重置界面 1-放大;0-还原

Private WithEvents mmessageManager As TimeLineMessageManager     '消息管理器
Attribute mmessageManager.VB_VarHelpID = -1
Private WithEvents mtimeLineControl As TimeLineControl           '时间轴控件
Attribute mtimeLineControl.VB_VarHelpID = -1

'常量定义
Private Const M_CON_TYPE_NORMAL As String = "普通文本"
Private Const M_CON_TYPE_LAYOUT As String = "排版文本"
Private Const M_CON_TYPE_GROUP As String = "分组"
Private Const M_CON_TYPE_COUNT As String = "计数数据"
Private Const M_CON_TYPE_DIVISION As String = "分割数据"
Private Const M_CON_TYPE_CONTINUOUS As String = "连续数据"
Private Const M_CON_TYPE_TICK As String = "时刻数据"
Private Const M_CON_TYPE_MEASURE As String = "标尺数据"
Private Const M_CON_TYPE_MEASUREVERTICALTEXT As String = "纵向文本"
Private Const M_CON_TYPE_CUSTOMTICK As String = "自定义时刻"
Private Const M_CON_TYPE_DATAAREA As String = "数据区域"

'常量索引
Private Const M_CON_KEY_用药跟踪_临 As String = "K_用药跟踪_临"
Private Const M_CON_KEY_用药跟踪_长 As String = "K_用药跟踪_长"
Private Const M_CON_KEY_病历 As String = "K_病历文件"
Private Const M_CON_KEY_检查 As String = "K_检查"
Private Const M_CON_KEY_检验 As String = "K_检验"
Private Const M_CON_KEY_其他长嘱 As String = "K_其他长嘱"
Private Const M_CON_KEY_其他临嘱 As String = "K_其他临嘱"
Private Const M_CON_KEY_手术 As String = "K_手术"
Private Const M_CON_KEY_住院天数 As String = "K_住院天数"
Private Const M_CON_KEY_手术后天数 As String = "K_手术后天数"
Private Const M_CON_KEY_纵向文本 As String = "K_纵向文本"

'颜色深度:浅红,淡红,鲜红,深红,暗红   与门诊一览保持一致
Private Enum CONST_COLOR
    '红色
    COLOR_浅红 = &HC0C0FF          '浅红
    COLOR_淡红 = &H8080FF
    COLOR_鲜红 = &HFF&
    COLOR_深红 = &HC0&
    COLOR_暗红 = &H80&
    '橙色
    COLOR_浅橙 = &HC0E0FF
    COLOR_淡橙 = &H80C0FF
    COLOR_鲜橙 = &H80FF&
    COLOR_深橙 = &H40C0&
    COLOR_暗橙 = &H4080&
    '黄色
    COLOR_浅黄 = &HC0FFFF
    COLOR_淡黄 = &H80FFFF
    COLOR_鲜黄 = &HFFFF&
    COLOR_深黄 = &HC0C0&
    COLOR_暗黄 = &H8080&
    '绿色
    COLOR_浅绿 = &HC0FFC0
    COLOR_淡绿 = &H80FF80
    COLOR_鲜绿 = &HFF00&
    COLOR_深绿 = &HC000&
    COLOR_暗绿 = &H8000&
    '青色
    COLOR_浅青 = &HFFFFC0
    COLOR_淡青 = &HFFFF80
    COLOR_鲜青 = &HFFFF00
    COLOR_深青 = &HC0C000
    COLOR_暗青 = &H808000
    '蓝色
    COLOR_浅蓝 = &HFFC0C0
    COLOR_淡蓝 = &HFF8080
    COLOR_鲜蓝 = &HFF0000
    COLOR_深蓝 = &HC00000
    COLOR_暗蓝 = &H800000
    '紫色
    COLOR_浅紫 = &HFFC0FF
    COLOR_淡紫 = &HFF80FF
    COLOR_鲜紫 = &HFF00FF
    COLOR_深紫 = &HC000C0
    COLOR_暗紫 = &H800080
    '白色
    COLOR_白色 = &H80000005
    COLOR_FORMBK = &H8000000B
    COLOR_CENTERBK = &H808000
End Enum

'鼠标按键
Private Enum CONST_MouseButtons
    MouseButtons_Left = &H100000
    MouseButtons_None = 0
    MouseButtons_Middle = &H400000
    MouseButtons_Right = &H200000
    MouseButtons_XButton1 = &H800000
    MouseButtons_XButton2 = &H1000000
End Enum

Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng科室ID As Long
Private mintBaby As Integer

Private mDatBegin As Date      '每一页的开始日期
Private mDatEnd As Date        '每一页的结束日期
Private mDatIn As Date       '记录入院日期
Private mdatOut As Date      '记录出院日期
Private mudtTimeLine As TimeLineData     '
Private mudtDesign As TimeLineDesignInfo
Private mbytFont As Byte           '0-小字体;1-大字体 显示字体风格
Private mlng医嘱ID As Long
Private mlng应用方式  As Long         '0-禁用;1-单独使用;2-心率与脉搏共用
Private mrsFrequency   As ADODB.Recordset     '护理项目频次   频次, 序号, 开始, 结束, 类别
Private mrs汇总时段 As ADODB.Recordset         '护理汇总时段    开始, 结束, 类别
Private mrs汇总项目 As ADODB.Recordset          '护理汇总项目

Private mobjPopup As CommandBarPopup     '分页按钮
Private mlngDay  As Long               '间隔天数(一页)取值范围默认 7天
Private mlngPages As Long              '总页数

'参数变量
Private mstr体温开始时间 As String    '参数号 =5 模块号 1255 控制标准体温单(不含专科)每天6个时点中首次开始的时点，6个时点间隔为4小时，如：参数值为4，6个时点分别为：4,8,12,16,20,24
Private mstr显示当天 As String   '参数号=42 模块号=1255  汇总波动显示当天数据
Private mblnMeasureArea As Boolean   'T-显示护理区域,F-隐藏护理区域


'-----------------------------------------------------------------------------------------------------------------------------
'刷新接口
Public Function zlRefresh(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng科室ID As Long, ByVal intBaby As Integer) As Boolean
'功能：
'参数:objFrmMain-主窗体
'lngPatiID-病人ID
'lngMainID-主页ID
'lng科室ID-科室ID
'intBaby-婴儿序号
    
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mlng科室ID = lng科室ID
    mintBaby = IIf(intBaby > 0, intBaby, 0)
    Call FuncLoadPages
End Function

Private Sub cbsSub_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim arrPara As Variant
    Dim lngPage As Long
    Dim i   As Long
    Dim udtDataItem As DataItem
    
    Select Case Control.ID
    
    Case conMenu_View_Jump
        arrPara = Split(Control.Parameter, ",")
        lngPage = Val(arrPara(0))
        mDatBegin = CDate(arrPara(1))
        mDatEnd = CDate(arrPara(2))
    Case conMenu_View_Forward, conMenu_View_Backward   '上一页,下一页
        arrPara = Split(mobjPopup.Parameter, ",")
        If Control.ID = conMenu_View_Forward Then
            lngPage = Val(arrPara(0)) - 1
            mDatBegin = mDatBegin - mlngDay
            mDatEnd = mDatEnd - mlngDay
        ElseIf Control.ID = conMenu_View_Backward Then
            lngPage = Val(arrPara(0)) + 1
            mDatBegin = mDatBegin + mlngDay
            mDatEnd = mDatEnd + mlngDay
        End If
    Case conMenu_Img_Look  '观片
        If mlng医嘱ID <> 0 Then
            RaiseEvent ViewPACSImage(mlng医嘱ID)
        End If
    Case conMenu_View_Show
        If Not mblnMeasureArea Then
            Control.Caption = "隐藏护理"
            Control.ToolTipText = "隐藏护理"
            Control.IconId = conMenu_Manage_Up
        Else
            Control.Caption = "显示护理"
            Control.ToolTipText = "显示护理"
            Control.IconId = conMenu_Manage_Down
        End If
        mblnMeasureArea = Not mblnMeasureArea
        Call SetFontSize(mbytFont)
    Case conMenu_Tool_Assistant
        Control.Checked = Not Control.Checked
        mtimeLineControl.IsShowReticle = Control.Checked
    Case conMenu_View_Navigatebeginning
        mtimeLineControl.ScrollToLeft
    Case conMenu_View_Navigateend
        mtimeLineControl.ScrollToRight
    Case conMenu_Process_Zoom
        Control.Checked = Not Control.Checked
        If Control.Checked Then
            Control.IconId = conMenu_Process_Small
            RaiseEvent ResizeForm(1)
            Control.Caption = "缩小"
            Control.ToolTipText = "缩小"
        Else
            Control.IconId = conMenu_Process_Zoom
            RaiseEvent ResizeForm(0)
            Control.Caption = "放大"
            Control.ToolTipText = "放大"
        End If
    End Select
    
    Select Case Control.ID
    Case conMenu_View_Jump, conMenu_View_Forward, conMenu_View_Backward
        mobjPopup.Caption = "第" & lngPage & "页：" & Format(mDatBegin, "YYYY-MM-DD") & "～" & Format(mDatEnd, "YYYY-MM-DD")
        mobjPopup.Parameter = lngPage & "," & Format(mDatBegin, "YYYY-MM-DD") & "," & Format(mDatEnd, "YYYY-MM-DD")
        mobjPopup.SetFocus
        For i = 1 To mobjPopup.CommandBar.Controls.Count
            If i = lngPage Then
                mobjPopup.CommandBar.Controls(i).Checked = True
            Else
                mobjPopup.CommandBar.Controls(i).Checked = False
            End If
        Next
        Call FuncCreateTimeLine
    End Select
End Sub

'---------------------------------------------------------------------------------------------------------------------------------
Private Sub cbsSub_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call cbsSub.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    On Error Resume Next
    With picMain
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - .Top
    End With
    mtimeLineControl.SetParentControl picMain.hwnd
End Sub

Private Sub cbsSub_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim arrPara As Variant
    
    arrPara = Split(mobjPopup.Parameter, ",")
    mobjPopup.Visible = mlngPages > 1
    
    Select Case Control.ID
    Case conMenu_View_Forward, conMenu_View_Backward   '上一页,下一页
        Control.Visible = mlngPages > 1
        If Control.Visible Then
            If Val(arrPara(0)) = 1 And Control.ID = conMenu_View_Forward Then
                Control.Enabled = False
            ElseIf Val(arrPara(0)) = mlngPages And Control.ID = conMenu_View_Backward Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        End If
    Case conMenu_View_Navigatebeginning, conMenu_View_Navigateend
        Control.Visible = (mDatEnd - mDatBegin) > 15
    End Select
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim tlbContext As TimeLineBusinessContext     '时间轴业务环境
    Dim strTmp As String
    
    '参数读取
    mstr体温开始时间 = zlDatabase.GetPara("体温开始时间", glngSys, p护理记录管理, "4")
    mstr显示当天 = zlDatabase.GetPara("汇总波动显示当天数据", glngSys, p护理记录管理, "0")
    
    Set tlbContext = New TimeLineBusinessContext
    Set mmessageManager = tlbContext.MessageManager

    Set mtimeLineControl = New TimeLineControl
    Set mtimeLineControl.BusinessContext = tlbContext
    mtimeLineControl.DockMode = TimeLineDockStyle_Fill '填满容器
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\显示护理", "显示护理", "1")
    mblnMeasureArea = IIf(strTmp = "1", True, False)
    
    Call InitCommandBar
    
    strSQL = "Select 频次, 序号, 开始, 结束, 类别 From 护理项目频次"
    Set mrsFrequency = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    strSQL = "Select 开始, 结束, 类别 From 护理汇总时段 Where 单据 = 1"
    Set mrs汇总时段 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    strSQL = "Select 序号, NVL(父序号,0) as 父序号 From 护理汇总项目"
    Set mrs汇总项目 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

End Sub

Private Sub Form_Resize()
    cbsSub_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.Visible Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\显示护理", "显示护理", IIf(mblnMeasureArea, "1", "0"))
    End If
    Set mtimeLineControl = Nothing
    Set mmessageManager = Nothing
End Sub

Private Function FuncMakeXMLDesign(ByRef udtDesign As TimeLineDesignInfo) As String
    Dim strDesign As String
    Dim strTmp As String
    Dim intDisplayVal As Integer, intTickStartTime As Integer
    Dim intTemp As Integer
    Dim i As Long

    Dim colTemp As Collection
    Dim udtTick  As DesignInfoTickRange
    
    strDesign = "<?xml version=""1.0"" encoding=""utf-16""?>" & vbCrLf & _
        "<TimeLineDesignInfo xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">" & vbCrLf
    With udtDesign
        Set colTemp = mudtTimeLine.colMeasureData
        If mbytFont = 1 Then
            .RowHeight = 22
            If colTemp.Count > 0 And colTemp.Count <= 3 Then
                .MeasureTitleWidth = 180 \ colTemp.Count
            Else
                .MeasureTitleWidth = 65
            End If
            .GridMinHeight = 220
            .DateTitleFont = Replace(.DateTitleFont, "9pt", "12pt")
            .TickTitleFont = Replace(.TickTitleFont, "9pt", "12pt")
            .Height = 200
            If mblnMeasureArea Then
                .TickWidth = 25
                .ShowTick = True
            Else
                .TickWidth = 150
                .ShowTick = False
            End If
        Else
            .RowHeight = 17
            If colTemp.Count > 0 And colTemp.Count <= 3 Then
                .MeasureTitleWidth = 150 \ colTemp.Count
            Else
                .MeasureTitleWidth = 50
            End If

            .GridMinHeight = 220
            .DateTitleFont = Replace(.DateTitleFont, "12pt", "9pt")
            .TickTitleFont = Replace(.TickTitleFont, "12pt", "9pt")
            .Height = 160
            If mblnMeasureArea Then
                .TickWidth = 20
                .ShowTick = True
            Else
                .TickWidth = 120
                .ShowTick = False
            End If
        End If

        strDesign = strDesign & _
                    IIf(.BackgroundColor <> Empty, Space(2) & "<BackgroundColor>" & .BackgroundColor & "</BackgroundColor>" & vbCrLf, "") & _
                    IIf(.DateTitle <> Empty, Space(2) & "<DateTitle>" & .DateTitle & "</DateTitle>" & vbCrLf, "") & _
                    IIf(.DateTitleFont <> Empty, Space(2) & "<DateTitleFont>" & .DateTitleFont & "</DateTitleFont>" & vbCrLf, "") & _
                    IIf(.DateTitleColor <> Empty, Space(2) & "<DateTitleColor>" & .DateTitleColor & "</DateTitleColor>" & vbCrLf, "") & _
                    IIf(.DateStart <> Empty, Space(2) & "<DateStart>" & .DateStart & "</DateStart>" & vbCrLf, "") & _
                    IIf(.dateEnd <> Empty, Space(2) & "<DateEnd>" & .dateEnd & "</DateEnd>" & vbCrLf, "") & _
                    IIf(.ShowTick <> Empty, Space(2) & "<ShowTick>" & IIf(.ShowTick, "true", "false") & "</ShowTick>" & vbCrLf, "") & _
                    IIf(.ShowFullDate <> Empty, Space(2) & "<ShowFullDate>" & IIf(.ShowFullDate, "true", "false") & "</ShowFullDate>" & vbCrLf, "") & _
                    IIf(.TickTitle <> Empty, Space(2) & "<TickTitle>" & .TickTitle & "</TickTitle>" & vbCrLf, "") & _
                    IIf(.TickTitleFont <> Empty, Space(2) & "<TickTitleFont>" & .TickTitleFont & "</TickTitleFont>" & vbCrLf, "") & _
                    IIf(.TickTitleColor <> Empty, Space(2) & "<TickTitleColor>" & .TickTitleColor & "</TickTitleColor>" & vbCrLf, "") & _
                    IIf(.TickWidth <> Empty, Space(2) & "<TickWidth>" & .TickWidth & "</TickWidth>" & vbCrLf, "")
        strDesign = strDesign & Space(2) & "<DesignInfoTickRangeList>" & vbCrLf
        strTmp = ""
        If .TickRangeListCount = 0 Then .TickRangeListCount = TICK_6   '缺省间隔为6
        intTemp = 24 \ .TickRangeListCount
        intDisplayVal = Val(mstr体温开始时间)
        intTickStartTime = 0
        For i = 1 To .TickRangeListCount
            strTmp = strTmp & _
                Space(4) & "<DesignInfoTickRange>" & vbCrLf & _
                    Space(6) & "<DisplayValue>" & intDisplayVal & "</DisplayValue>" & vbCrLf & _
                    Space(6) & "<TickStartTime>" & intTickStartTime & ":0" & "</TickStartTime>" & vbCrLf & _
                Space(4) & "</DesignInfoTickRange>" & vbCrLf
            intDisplayVal = intDisplayVal + intTemp  '下一个显示值
            intTickStartTime = intTickStartTime + intTemp
        Next
    
        strDesign = strDesign & strTmp & "</DesignInfoTickRangeList>" & vbCrLf
        '字体更换
        If mbytFont = 1 Then
            If .DateFont <> Empty Then .DateFont = Replace(.DateFont, "9pt", "12pt")
            If .TickFont <> Empty Then .TickFont = Replace(.TickFont, "9pt", "12pt")
        Else
            If .DateFont <> Empty Then .DateFont = Replace(.DateFont, "12pt", "9pt")
            If .TickFont <> Empty Then .TickFont = Replace(.TickFont, "12pt", "9pt")
        End If
        strDesign = strDesign & _
                    IIf(.DateFont <> Empty, Space(2) & "<DateFont>" & .DateFont & "</DateFont>" & vbCrLf, "") & _
                    IIf(.TickFont <> Empty, Space(2) & "<TickFont>" & .TickFont & "</TickFont>" & vbCrLf, "") & _
                    IIf(.MergePeriodWidth <> Empty, Space(2) & "<MergePeriodWidth>" & .MergePeriodWidth & "</MergePeriodWidth>" & vbCrLf, "") & _
                    IIf(.EmptyDataMergeDayCount <> Empty, Space(2) & "<EmptyDataMergeDayCount>" & .EmptyDataMergeDayCount & "</EmptyDataMergeDayCount>" & vbCrLf, "") & _
                    IIf(.EmptyDataMergePeriodWidth <> Empty, Space(2) & "<EmptyDataMergePeriodWidth>" & .EmptyDataMergePeriodWidth & "</EmptyDataMergePeriodWidth>" & vbCrLf, "") & _
                    IIf(.PaddingLeft <> Empty, Space(2) & "<PaddingLeft>" & .PaddingLeft & "</PaddingLeft>" & vbCrLf, "") & _
                    IIf(.PaddingTop <> Empty, Space(2) & "<PaddingTop>" & .PaddingTop & "</PaddingTop>" & vbCrLf, "") & _
                    IIf(.PaddingRight <> Empty, Space(2) & "<PaddingRight>" & .PaddingRight & "</PaddingRight>" & vbCrLf, "") & _
                    IIf(.PaddingBottom <> Empty, Space(2) & "<PaddingBottom>" & .PaddingBottom & "</PaddingBottom>" & vbCrLf, "") & _
                    IIf(.RowHeight <> Empty, Space(2) & "<RowHeight>" & .RowHeight & "</RowHeight>" & vbCrLf, "")
    
        strDesign = strDesign & _
                    Space(2) & "<Measure>" & vbCrLf & _
                        IIf(.MeasureTitleWidth <> Empty, Space(4) & "<MeasureTitleWidth>" & .MeasureTitleWidth & "</MeasureTitleWidth>" & vbCrLf, "") & _
                        IIf(.GridMinHeight <> Empty, Space(4) & "<GridMinHeight>" & .GridMinHeight & "</GridMinHeight>" & vbCrLf, "") & _
                        IIf(.TopFixedSmallRowCount <> Empty, Space(4) & "<TopFixedSmallRowCount>" & .TopFixedSmallRowCount & "</TopFixedSmallRowCount>" & vbCrLf, "") & _
                        IIf(.BottomFixedSmallRowCount <> Empty, Space(4) & "<BottomFixedSmallRowCount>" & .BottomFixedSmallRowCount & "</BottomFixedSmallRowCount>" & vbCrLf, "") & _
                        IIf(.GridYSplitCount <> Empty, Space(4) & "<GridYSplitCount>" & .GridYSplitCount & "</GridYSplitCount>" & vbCrLf, "") & _
                        IIf(.GridYSmallSplitCount <> Empty, Space(4) & "<GridYSmallSplitCount>" & .GridYSmallSplitCount & "</GridYSmallSplitCount>" & vbCrLf, "") & _
                        IIf(.Height <> Empty, Space(4) & "<Height>" & .Height & "</Height>" & vbCrLf, "") & _
                    Space(2) & "</Measure>" & vbCrLf

    End With
    strDesign = strDesign & "</TimeLineDesignInfo>"
    
    FuncMakeXMLDesign = strDesign
End Function

Private Function FuncMakeXMLDataList(ByRef colList As Collection, Optional ByVal intSpace As Integer) As String
'功能:
'colList-DataItem\DataInfo的集合
    Dim udtItem As DataItem
    Dim udtInfo As DataInfo
    Dim varItem As Variant
    Dim strRet As String
    Dim strHead As String
    Dim strFoot As String
    
    If colList Is Nothing Then Exit Function

    If colList.Count = 0 Then Exit Function

    If TypeName(colList(colList.Count)) = TypeName(udtItem) Then
        strHead = Space(intSpace) & "<ListDataItem>" & vbCrLf
        strFoot = Space(intSpace) & "</ListDataItem>" & vbCrLf
    ElseIf TypeName(colList(colList.Count)) = TypeName(udtInfo) Then
        strHead = Space(intSpace) & "<ListDataInfo>" & vbCrLf
        strFoot = Space(intSpace) & "</ListDataInfo>" & vbCrLf
    End If
    For Each varItem In colList
        strRet = strRet & FuncMakeXMLData(varItem, (intSpace + 2))
    Next
    FuncMakeXMLDataList = strHead & strRet & strFoot
End Function

Private Function FuncMakeXMLData(ByRef varData As Variant, Optional ByVal intSpace As Integer) As String
    Dim strRet As String
    Dim udtItem As DataItem
    Dim udtInfo As DataInfo

    If TypeName(varData) = TypeName(udtItem) Then
        udtItem = varData
        With udtItem
            If mbytFont = 1 Then
                If .Font <> Empty Then .Font = Replace(.Font, "9pt", "12pt")
                If .HotspotFont <> Empty Then .HotspotFont = Replace(.HotspotFont, "9pt", "12pt")
                If .TitleFont <> Empty Then .TitleFont = Replace(.TitleFont, "9pt", "12pt")
            Else
                If .Font <> Empty Then .Font = Replace(.Font, "12pt", "9pt")
                If .HotspotFont <> Empty Then .HotspotFont = Replace(.HotspotFont, "12pt", "9pt")
                If .TitleFont <> Empty Then .TitleFont = Replace(.TitleFont, "12pt", "9pt")
            End If
            strRet = Space(intSpace + 2) & "<DataItem>" & vbCrLf
                '公有属性
                strRet = strRet & IIf(.GraphType <> Empty, Space(intSpace + 4) & "<GraphType>" & .GraphType & "</GraphType>" & vbCrLf, "")
                strRet = strRet & IIf(.Title <> Empty, Space(intSpace + 4) & "<Title>" & .Title & "</Title>" & vbCrLf, "")
                strRet = strRet & IIf(.TextColor <> Empty, Space(intSpace + 4) & "<TextColor>" & .TextColor & "</TextColor>" & vbCrLf, "")
                strRet = strRet & IIf(.Color <> Empty, Space(intSpace + 4) & "<Color>" & .Color & "</Color>" & vbCrLf, "")
                strRet = strRet & IIf(.Font <> Empty, Space(intSpace + 4) & "<Font>" & .Font & "</Font>" & vbCrLf, "")
                strRet = strRet & IIf(.BackgroundColor <> Empty, Space(intSpace + 4) & "<BackgroundColor>" & .BackgroundColor & "</BackgroundColor>" & vbCrLf, "")
                strRet = strRet & IIf(.ShowHotspotEffect = True, Space(intSpace + 4) & "<ShowHotspotEffect>true</ShowHotspotEffect>" & vbCrLf, "")
                strRet = strRet & IIf(.HotspotFont <> Empty, Space(intSpace + 4) & "<HotspotFont>" & .HotspotFont & "</HotspotFont>" & vbCrLf, "")
                strRet = strRet & IIf(.HotspotColor <> Empty, Space(intSpace + 4) & "<HotspotColor>" & .HotspotColor & "</HotspotColor>" & vbCrLf, "")
                strRet = strRet & IIf(.ShowHotspotCursor = True, Space(intSpace + 4) & "<ShowHotspotCursor>true</ShowHotspotCursor>" & vbCrLf, "")
                strRet = strRet & IIf(.TitleFont <> Empty, Space(intSpace + 4) & "<TitleFont>" & .TitleFont & "</TitleFont>" & vbCrLf, "")
                strRet = strRet & IIf(.TitleColor <> Empty, Space(intSpace + 4) & "<TitleColor>" & .TitleColor & "</TitleColor>" & vbCrLf, "")
                strRet = strRet & IIf(.GroupPosition <> Empty, Space(intSpace + 4) & "<GroupPosition>" & .GroupPosition & "</GroupPosition>" & vbCrLf, "")
                '私有属性
                Select Case .GraphType
                
                Case M_CON_TYPE_GROUP   '分组
                    'BackgroundColor,Title,GraphType,GroupPosition
                Case M_CON_TYPE_COUNT, M_CON_TYPE_CONTINUOUS, M_CON_TYPE_NORMAL, M_CON_TYPE_TICK, M_CON_TYPE_LAYOUT '计数数据,连续数据,普通文本,时刻数据,排版文本
                    'GraphType,Title,BackgroundColor,TextColor
                    If .GraphType = M_CON_TYPE_LAYOUT Then
                        strRet = strRet & IIf(.BorderColor <> Empty, Space(intSpace + 4) & "<BorderColor>" & .BorderColor & "</BorderColor>" & vbCrLf, "")
                    ElseIf .GraphType = M_CON_TYPE_CONTINUOUS Then
                        strRet = strRet & IIf(.Effect <> Empty, Space(intSpace + 4) & "<Effect>" & .Effect & "</Effect>" & vbCrLf, "")
                    End If
                Case M_CON_TYPE_DIVISION '分割数据
                    'Font,TextColor,BackgroundColor,Title
                     strRet = strRet & IIf(.SplitString <> Empty, Space(intSpace + 4) & "<SplitString>" & .SplitString & "</SplitString>" & vbCrLf, "") & _
                    IIf(.SplitCount <> Empty, Space(intSpace + 4) & "<SplitCount>" & .SplitCount & "</SplitCount>" & vbCrLf, "")
                Case M_CON_TYPE_MEASURE      '标尺数据
                    'Color,Title
                    strRet = strRet & IIf(.Unit <> Empty, Space(intSpace + 4) & "<Unit>" & .Unit & "</Unit>" & vbCrLf, "") & _
                    IIf(.MinValue <> Empty, Space(intSpace + 4) & "<MinValue>" & .MinValue & "</MinValue>" & vbCrLf, "") & _
                    IIf(.MaxValue <> Empty, Space(intSpace + 4) & "<MaxValue>" & .MaxValue & "</MaxValue>" & vbCrLf, "") & _
                    IIf(.SplitNum <> Empty, Space(intSpace + 4) & "<SplitNum>" & .SplitNum & "</SplitNum>" & vbCrLf, "") & _
                    IIf(.SplitScale <> Empty, Space(intSpace + 4) & "<SplitScale>" & .SplitScale & "</SplitScale>" & vbCrLf, "") & _
                    IIf(.IsDataDynamicExpansion <> Empty, Space(intSpace + 4) & "<IsDataDynamicExpansion>" & IIf(.IsDataDynamicExpansion, "true", "false") & "</IsDataDynamicExpansion>" & vbCrLf, "") & _
                    IIf(.ShadowTitle <> Empty, Space(intSpace + 4) & "<ShadowTitle>" & .ShadowTitle & "</ShadowTitle>" & vbCrLf, "") & _
                    IIf(.ShadowColor <> Empty, Space(intSpace + 4) & "<ShadowColor>" & .ShadowColor & "</ShadowColor>" & vbCrLf, "") & _
                    IIf(.BalloonColor <> Empty, Space(intSpace + 4) & "<BalloonColor>" & .BalloonColor & "</BalloonColor>" & vbCrLf, "") & _
                    IIf(.BalloonTitle <> Empty, Space(intSpace + 4) & "<BalloonTitle>" & .BalloonTitle & "</BalloonTitle>" & vbCrLf, "") & _
                    IIf(.LegendType <> Empty, Space(intSpace + 4) & "<LegendType>" & .LegendType & "</LegendType>" & vbCrLf, "") & _
                    IIf(.ShadowLegendType <> Empty, Space(intSpace + 4) & "<ShadowLegendType>" & .ShadowLegendType & "</ShadowLegendType>" & vbCrLf, "") & _
                    IIf(.BalloonLegendType <> Empty, Space(intSpace + 4) & "<BalloonLegendType>" & .BalloonLegendType & "</BalloonLegendType>" & vbCrLf, "")
                Case M_CON_TYPE_MEASUREVERTICALTEXT    '纵向文本

                Case M_CON_TYPE_CUSTOMTICK   '自定义时刻
                    strRet = strRet & IIf(.StartDate <> Empty, Space(intSpace + 4) & "<StartDate>" & .StartDate & "</StartDate>" & vbCrLf, "") & _
                    IIf(.EndDate <> Empty, Space(intSpace + 4) & "<EndDate>" & .EndDate & "</EndDate>" & vbCrLf, "") & _
                    IIf(.FixedTick <> Empty, Space(intSpace + 4) & "<FixedTick>" & .FixedTick & "</FixedTick>" & vbCrLf, "") & _
                    IIf(.EquantTick <> Empty, Space(intSpace + 4) & "<EquantTick>" & .EquantTick & "</EquantTick>" & vbCrLf, "") & _
                    IIf(.EquantTickUnit <> Empty, Space(intSpace + 4) & "<EquantTickUnit>" & .EquantTickUnit & "</EquantTickUnit>" & vbCrLf, "") & _
                    IIf(.TickWidth <> Empty, Space(intSpace + 4) & "<TickWidth>" & .TickWidth & "</TickWidth>" & vbCrLf, "")
                Case M_CON_TYPE_DATAAREA     '数据区域
                    strRet = strRet & IIf(.LineColor <> Empty, Space(intSpace + 4) & "<LineColor>" & .LineColor & "</LineColor>" & vbCrLf, "") & _
                    IIf(.StartDate <> Empty, Space(intSpace + 4) & "<StartDate>" & .StartDate & "</StartDate>" & vbCrLf, "") & _
                    IIf(.EndDate <> Empty, Space(intSpace + 4) & "<EndDate>" & .EndDate & "</EndDate>" & vbCrLf, "") & _
                    IIf(.IsCollapse <> Empty, Space(intSpace + 4) & "<IsCollapse>" & .IsCollapse & "</IsCollapse>" & vbCrLf, "")
                End Select
                
                If Not .ListData Is Nothing Then
                    If .ListData.Count > 0 Then
                        strRet = strRet & FuncMakeXMLDataList(.ListData, intSpace + 4)
                    End If
                End If
                
            strRet = strRet & Space(intSpace + 2) & "</DataItem>" & vbCrLf
    
        End With
    ElseIf TypeName(varData) = TypeName(udtInfo) Then
        udtInfo = varData
        With udtInfo
            If mbytFont = 1 Then
                If .Font <> Empty Then .Font = Replace(.Font, "9pt", "12pt")
                If .HotspotFont <> Empty Then .HotspotFont = Replace(.HotspotFont, "9pt", "12pt")
            Else
                If .Font <> Empty Then .Font = Replace(.Font, "12pt", "9pt")
                If .HotspotFont <> Empty Then .HotspotFont = Replace(.HotspotFont, "12pt", "9pt")
            End If
            strRet = Space(intSpace + 2) & "<DataInfo>" & vbCrLf
            strRet = strRet & IIf(.Value <> Empty, Space(intSpace + 4) & "<Value>" & .Value & "</Value>" & vbCrLf, "")
            strRet = strRet & IIf(.Time <> Empty, Space(intSpace + 4) & "<Time>" & .Time & "</Time>" & vbCrLf, "")
            strRet = strRet & IIf(.RowNumber <> Empty, Space(intSpace + 4) & "<RowNumber>" & .RowNumber & "</RowNumber>" & vbCrLf, "")
            strRet = strRet & IIf(.TimeEnd <> Empty, Space(intSpace + 4) & "<TimeEnd>" & .TimeEnd & "</TimeEnd>" & vbCrLf, "")
            strRet = strRet & IIf(.Tag <> Empty, Space(intSpace + 4) & "<Tag>" & .Tag & "</Tag>" & vbCrLf, "")
            strRet = strRet & IIf(.BackgroundColor <> Empty, Space(intSpace + 4) & "<BackgroundColor>" & .BackgroundColor & "</BackgroundColor>" & vbCrLf, "")
            strRet = strRet & IIf(.TextColor <> Empty, Space(intSpace + 4) & "<TextColor>" & .TextColor & "</TextColor>" & vbCrLf, "")
            strRet = strRet & IIf(.Font <> Empty, Space(intSpace + 4) & "<Font>" & .Font & "</Font>" & vbCrLf, "")
            strRet = strRet & IIf(.HotspotFont <> Empty, Space(intSpace + 4) & "<HotspotFont>" & .HotspotFont & "</HotspotFont>" & vbCrLf, "")
            strRet = strRet & IIf(.ShowHotspotCursor = True, Space(intSpace + 4) & "<ShowHotspotCursor>true</ShowHotspotCursor>" & vbCrLf, "")
            strRet = strRet & IIf(.HotspotColor <> Empty, Space(intSpace + 4) & "<HotspotColor>" & .HotspotColor & "</HotspotColor>" & vbCrLf, "")
            strRet = strRet & IIf(.RowIndex <> Empty, Space(intSpace + 4) & "<RowIndex>" & .RowIndex & "</RowIndex>" & vbCrLf, "")
            strRet = strRet & IIf(.LegendType <> Empty, Space(intSpace + 4) & "<LegendType>" & .LegendType & "</LegendType>" & vbCrLf, "")
            strRet = strRet & IIf(.ShadowLegendType <> Empty, Space(intSpace + 4) & "<ShadowLegendType>" & .ShadowLegendType & "</ShadowLegendType>" & vbCrLf, "")
            strRet = strRet & IIf(.BalloonLegendType <> Empty, Space(intSpace + 4) & "<BalloonLegendType>" & .BalloonLegendType & "</BalloonLegendType>" & vbCrLf, "")
            strRet = strRet & IIf(.NumberValue <> Empty, Space(intSpace + 4) & "<NumberValue>" & .NumberValue & "</NumberValue>" & vbCrLf, "")
            strRet = strRet & IIf(.ShadowValue <> Empty, Space(intSpace + 4) & "<ShadowValue>" & .ShadowValue & "</ShadowValue>" & vbCrLf, "")
            strRet = strRet & IIf(.BalloonValue <> Empty, Space(intSpace + 4) & "<BalloonValue>" & .BalloonValue & "</BalloonValue>" & vbCrLf, "")
            strRet = strRet & IIf(.Tip <> Empty, Space(intSpace + 4) & "<Tip>" & .Tip & "</Tip>" & vbCrLf, "")
            strRet = strRet & IIf(.Group <> Empty, Space(intSpace + 4) & "<Group>" & .Group & "</Group>" & vbCrLf, "")
            strRet = strRet & Space(intSpace + 2) & "</DataInfo>" & vbCrLf
        End With
    End If
    FuncMakeXMLData = strRet
End Function

Private Function FuncMakeXMLTimeLine(udtData As TimeLineData) As String
    Dim strRet As String
    Dim strTmp As String
    Dim varItem As Variant
    Dim udtDataItem As DataItem
    Dim colTmp As Collection
    Dim objXML As zl9ComLib.clsXML
    Dim lngBegin As Long, lngEnd As Long

    strRet = "<?xml version=""1.0"" encoding=""utf-16""?>" & vbCrLf & _
            "<TimeLineData xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">" & vbCrLf
    
    '表头数据
    If Not udtData.colHeaderData Is Nothing Then
        strRet = strRet & _
            Space(2) & "<HeaderData>" & vbCrLf & _
            Space(4) & "<ListDataItem>" & vbCrLf
        For Each varItem In udtData.colHeaderData
            strRet = strRet & FuncMakeXMLData(varItem, 4)
        Next
        strRet = strRet & Space(4) & "</ListDataItem>" & vbCrLf & _
            Space(2) & "</HeaderData>" & vbCrLf
    End If
    '页脚数据
    If Not udtData.colFooterData Is Nothing Then
        strRet = strRet & _
            Space(2) & "<FooterData>" & vbCrLf & _
            Space(4) & "<ListDataItem>" & vbCrLf
        For Each varItem In udtData.colFooterData
            If varItem.Title = "护理项目" And mblnMeasureArea = False Then
                '隐藏护理项目
            Else
                strRet = strRet & FuncMakeXMLData(varItem, 4)
            End If
            
        Next
        strRet = strRet & Space(4) & "</ListDataItem>" & vbCrLf & _
            Space(2) & "</FooterData>" & vbCrLf
    End If

    '标尺数据
    If Not udtData.colMeasureData Is Nothing Then
        strRet = strRet & _
            Space(2) & "<MeasureData>" & vbCrLf & _
            Space(4) & "<ListDataItem>" & vbCrLf
        For Each varItem In udtData.colMeasureData
            strRet = strRet & FuncMakeXMLData(varItem, 4)
        Next
        strRet = strRet & Space(4) & "</ListDataItem>" & vbCrLf & _
            Space(2) & "</MeasureData>" & vbCrLf
    End If

    '纵向文本数据
    If Not udtData.colMeasureVerticalText Is Nothing Then
        strRet = strRet & _
            Space(2) & "<MeasureVerticalText>" & vbCrLf & _
                Space(4) & "<ListDataItem>" & vbCrLf
        For Each varItem In udtData.colMeasureVerticalText
            strRet = strRet & FuncMakeXMLData(varItem, 4)
        Next
        strRet = strRet & Space(4) & "</ListDataItem>" & vbCrLf & _
            Space(2) & "</MeasureVerticalText>" & vbCrLf
    End If

    '自定义时刻
    If Not udtData.colCustomTick Is Nothing Then
        strRet = strRet & _
            Space(2) & "<CustomTick>" & vbCrLf & _
                Space(4) & "<ListDataItem>" & vbCrLf
        For Each varItem In udtData.colCustomTick
            strRet = strRet & FuncMakeXMLData(varItem, 4)
        Next
        strRet = strRet & Space(4) & "</ListDataItem>" & vbCrLf & _
            Space(2) & "</CustomTick>" & vbCrLf
    End If
    '数据区域
    If Not udtData.colDataArea Is Nothing Then
        strRet = strRet & _
            Space(2) & "<DataArea>" & vbCrLf & _
            Space(4) & "<ListDataItem>" & vbCrLf
        For Each varItem In udtData.colDataArea
            strRet = strRet & FuncMakeXMLData(varItem, 4)
        Next
        strRet = strRet & Space(4) & "</ListDataItem>" & vbCrLf & _
            Space(2) & "</DataArea>" & vbCrLf
    End If
    
    strRet = strRet & "</TimeLineData>"

    FuncMakeXMLTimeLine = strRet
End Function

Private Sub FuncClearUDT(varItem As Variant)
    Dim udtItem As DataItem
    Dim udtInfo As DataInfo
    
    Select Case TypeName(varItem)
    Case TypeName(udtItem)
        With varItem
            .BackgroundColor = Empty
            .BalloonLegendType = Empty
            .BalloonTitle = Empty
            .Color = Empty
            .EndDate = Empty
            .EquantTick = Empty
            .EquantTickUnit = Empty
            .FixedTick = Empty
            .Font = Empty
            .ShowHotspotEffect = False
            .HotspotFont = Empty
            .HotspotColor = Empty
            .GraphType = Empty
            .IsCollapse = Empty
            .IsDataDynamicExpansion = Empty
            .LegendType = Empty
            .LineColor = Empty
            Set .ListData = Nothing
            .MaxValue = Empty
            .MinValue = Empty
            .ShadowLegendType = Empty
            .ShadowTitle = Empty
            .SplitCount = Empty
            .SplitNum = Empty
            .SplitScale = Empty
            .SplitString = Empty
            .StartDate = Empty
            .TextColor = Empty
            .TickWidth = Empty
            .Title = Empty
            .Unit = Empty
            .BorderColor = Empty
            .ShowHotspotCursor = False
            .TitleColor = Empty
            .TitleFont = Empty
            .ShadowColor = Empty
            .BalloonColor = Empty
            .GroupPosition = Empty
            .Effect = Empty
            
            '
            .ItemTag = Empty
        End With
    Case TypeName(udtInfo)
        With varItem
            .BackgroundColor = Empty
            .BalloonLegendType = Empty
            .BalloonValue = Empty
            .Font = Empty
            .HotspotFont = Empty
            .HotspotColor = Empty
            .LegendType = Empty
            .NumberValue = Empty
            .RowIndex = Empty
            .RowNumber = Empty
            .ShadowLegendType = Empty
            .ShadowValue = Empty
            .Tag = Empty
            .TextColor = Empty
            .Time = Empty
            .TimeEnd = Empty
            .Value = Empty
            .Tip = Empty
            .Group = Empty
            .ShowHotspotCursor = False
        End With
    End Select
End Sub

Private Sub mmessageManager_ErrorShow(ByVal sender As Variant, ByVal e As ZLSoft_BusinessHome_ClientControl_TimeLineBase.IEventArgsMessageInfo)
    Dim objE As Object          '解决Win Server 2003\XP 下未安装.NET环境生成部件提示加载DLL错误
    Set objE = e
    LogWrite "住院一览的调试日志", "" & glngModul, "mmessageManager_ErrorShow", "mmessageManager_ErrorShow:" & objE.Caption & vbCrLf & objE.Message & vbCrLf & objE.Exception
End Sub

Private Sub mmessageManager_InfoShow(ByVal sender As Variant, ByVal e As ZLSoft_BusinessHome_ClientControl_TimeLineBase.IEventArgsMessageInfo)
    Dim objE As Object          '解决Win Server 2003\XP 下未安装.NET环境生成部件提示加载DLL错误
    Set objE = e
    LogWrite "住院一览的调试日志", "" & glngModul, "mmessageManager_InfoShow", "mmessageManager_InfoShow:" & objE.Caption & vbCrLf & objE.Message & vbCrLf & objE.Exception
End Sub

Private Sub mtimeLineControl_DataMouseClick(ByVal sender As Variant, ByVal e As ZLSoft_BusinessHome_ClientControl_TimeLineBase.IEventArgsDataInfo)
    Dim objPopup As CommandBarPopup
    Dim strTmp As String, strType As String
    Dim strMsg As String, str检查报告ID As String
    Dim lng报告ID As Long, lng医嘱ID As Long
    Dim objE As Object          '解决Win Server 2003\XP 下未安装.NET环境生成部件提示加载DLL错误
    Set objE = e
    If objE.MouseButtons = MouseButtons_Left Then
        Select Case objE.Name
        Case "病历文书"
            strTmp = objE.Tag
            If strTmp = "" Then Exit Sub
            If Len(strTmp) < 32 Then
                lng报告ID = CLng(strTmp) '老版病历查看
                Call gobjRichEPR.ViewDocument(Me, lng报告ID, False)
            ElseIf Len(strTmp) = 32 And Not gobjEmr Is Nothing Then
                '新版病历
                On Error Resume Next
                strMsg = gobjEmr.OpenInEPR(strTmp)
                err.Clear: On Error GoTo 0
            End If
         Case "检查", "检验"
            '查阅报告
            strTmp = objE.Tag
            strType = Split(strTmp, ",")(0)
            lng报告ID = Val(Split(strTmp, ",")(1))
            lng医嘱ID = Val(objE.RowIndex)
            If objE.Name = "检查" Then
                str检查报告ID = Split(strTmp, ",")(2)
                Call FuncEPRReport(Me, lng医嘱ID, "D", lng报告ID, str检查报告ID, 2)
            ElseIf objE.Name = "检验" Then
                Call FuncEPRReport(Me, lng医嘱ID, "", lng报告ID, , 2)
            End If
        End Select
    End If
    
    If objE.MouseButtons = MouseButtons_Right Then
        If objE.Name = "检查" Then
            mlng医嘱ID = Val(objE.RowIndex)
            Set objPopup = cbsSub.ActiveMenuBar.FindControl(, conMenu_EditPopup)
            If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Public Sub FuncCreateTimeLine()
'功能：加载数据
'参数:bytFunc =1   缺省样式
'     bytFunc =2   数据更新
    Dim strTitle As String
    Dim strDesignInfo As String
    Dim strBegin As String, strEnd As String
    Dim strSQL As String, strKey As String
    Dim i As Long, j As Long, lngPosition As Long
    Dim rsNurse As ADODB.Recordset       '体温记录项目
    Dim rsTemp As ADODB.Recordset
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    
    Dim colTmp As Collection, colHead As Collection
    Dim colMeaSure As Collection
    Dim colFoot As Collection
    Dim colItem As Collection
    
    Dim strFile As String
    Dim strData As String
    Dim datBegin As Date, datEnd As Date
    
    On Error GoTo errH
    strEnd = Format(mDatEnd, "YYYY-MM-DDTHH:MM:SS")
    strBegin = Format(mDatBegin, "YYYY-MM-DDTHH:MM:SS")
    
    datBegin = CDate(Format(mDatBegin, "YYYY-MM-DD 00:00:00"))
    datEnd = CDate(Format(mDatEnd, "YYYY-MM-DD 23:59:59"))
    '初始化时间轴对象
    '体温记录项目.记录频次  为NULL时默认为2
    strSQL = "Select a.项目序号, a.排列序号, a.记录名, a.记录法, a.记录符, a.记录色, a.最大值, a.最小值, a.单位值, Nvl(a.单位, b.项目单位) As 单位," & vbNewLine & _
            " NVL(a.记录频次,2) as 记录频次, a.刻度间隔, b.项目性质, b.应用方式, Decode(c.项目序号, Null, Decode(d.序号, Null, 0, 2), 1) As 项目类型, Nvl(d.父序号, 0) As 父序号 " & vbNewLine & _
            "From 体温记录项目 A, 护理记录项目 B, 护理波动项目 C, 护理汇总项目 D " & vbNewLine & _
            "Where a.项目序号 = b.项目序号 And b.项目序号 = c.项目序号(+) And b.项目序号 = d.序号(+) And (b.项目性质 = 1 Or b.项目性质 = 2 And Exists" & vbNewLine & _
            "       (Select 1" & vbNewLine & _
            "       From 病人护理文件 A, 病历文件列表 E, 病人护理数据 C, 病人护理明细 D" & vbNewLine & _
            "       Where a.格式id = e.Id And a.Id = c.文件id And c.Id = d.记录id And a.病人id = [1] And a.主页id = [2] And" & vbNewLine & _
            "       Nvl(a.婴儿, 0) = [3] And d.项目序号 = b.项目序号 And e.种类 = 3 And e.保留 = -1 And" & vbNewLine & _
            "       c.发生时间 Between [4] And [5])) And b.应用方式 > 0 And (b.适用科室 = 1 Or (b.适用科室 = 2 And Exists (Select 1 From 护理适用科室 Where 科室id = [6]))) And Instr([7], ',' || b.适用病人 || ',') > 0" & vbNewLine & _
            "Order By 记录法, 排列序号, 项目序号"
            
    Set rsNurse = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, mintBaby, datBegin, datEnd, mlng科室ID, IIf(mintBaby > 0, ",0,2,", ",0,1,"))
    
    With mudtDesign
        .BackgroundColor = "200,255,255,255"
        .DateTitle = "日期"
        .DateTitleFont = "宋体,9pt"
        .DateStart = strBegin
        .dateEnd = strEnd
        .ShowTick = IIf(mblnMeasureArea, True, False)
        .ShowFullDate = True
        .TickTitle = "时刻"
        .TickTitleFont = "宋体,9pt"
        .TickWidth = 20
        .TickRangeListCount = TICK_6
        .DateFont = "宋体,9pt"
        .TickFont = "宋体,9pt"
        .MergePeriodWidth = 50
        .EmptyDataMergeDayCount = 10    '超过9列无数据允许合并
        .EmptyDataMergePeriodWidth = 50
        .PaddingLeft = 10
        .PaddingTop = 10
        .PaddingRight = 10
        .PaddingBottom = 10
        .RowHeight = 22
        .MeasureTitleWidth = 60
        .GridMinHeight = 220
        .TopFixedSmallRowCount = 1
        .BottomFixedSmallRowCount = 1
        .GridYSplitCount = 5
        .GridYSmallSplitCount = 1
        .Height = 400
    End With
    '住院天数---------------------------------------------
    Set colHead = New Collection
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_NORMAL
        .Title = "住院天数"
        .TitleFont = "宋体,9pt"
        .TextColor = "Black"
        .Font = "宋体,9pt"
        Set .ListData = colTmp
    End With
    colHead.Add udtDataItem, M_CON_KEY_住院天数
    '手术后天数 计数数据---------------------------------------------
    Call FuncClearUDT(udtDataItem)
    With udtDataItem
        .GraphType = M_CON_TYPE_COUNT
        .Title = "手术后天数"
        .TitleFont = "宋体,9pt"
        .TextColor = "Black"
        .Font = "宋体,9pt"
        Set .ListData = colTmp
    End With
    colHead.Add udtDataItem, M_CON_KEY_手术后天数
    'MeasureData 标尺---------------------------------------------
    Set colMeaSure = New Collection
    '心率 根据护理记录项目设置情况来决定是否显示心率
    rsNurse.Filter = "项目序号=-1"
    If Not rsNurse.EOF Then mlng应用方式 = Nvl(rsNurse!应用方式, 0)

    rsNurse.Filter = "记录法=1 And 应用方式=1" '标尺数据 心率\脉搏\体温
    For i = 1 To rsNurse.RecordCount
        FuncClearUDT udtDataItem
        With udtDataItem
            .GraphType = M_CON_TYPE_MEASURE
            .Title = rsNurse!记录名
            .TitleFont = "宋体,9pt"
            .TextColor = "Black"
            .Unit = rsNurse!单位 & ""
            .MinValue = rsNurse!最小值
            .MaxValue = rsNurse!最大值
            j = (rsNurse!最大值 - rsNurse!最小值) / Nvl(rsNurse!刻度间隔, 1)
            If rsNurse!记录名 = "体温" Then
                j = 5
            ElseIf rsNurse!记录名 = "心率" Or rsNurse!记录名 = "脉搏" Then
                j = 7
            Else
                j = (rsNurse!最大值 - rsNurse!最小值) / Nvl(rsNurse!刻度间隔, 1)
            End If
            .SplitNum = j
            .Color = FuncColorRGB(CLng(rsNurse!记录色 & ""))
            .LegendType = "实心圆"  '（实心圆、空心圆、粗空心圆、点、叉、H符号）
            .ShadowLegendType = "空心圆"     '阴影点
            .BalloonLegendType = "空心圆"    '气球点
            .IsDataDynamicExpansion = True
            If rsNurse!项目序号 = 1 Then
            '体温
                .BalloonTitle = "物理升降温"
                .BalloonColor = "Red"
            ElseIf rsNurse!项目序号 = 2 Then
            '脉搏
                .ShadowTitle = "心率"
                If mlng应用方式 = 1 Or mlng应用方式 = 2 Then
                    lngPosition = rsNurse.AbsolutePosition
                    '心率单独应用时;脉搏的阴影区域颜色取心率的颜色值
                    rsNurse.Filter = "项目序号=-1"
                    If Not rsNurse.EOF Then
                        .ShadowColor = FuncColorRGB(CLng(rsNurse!记录色 & ""))
                    End If
                    '恢复到原来指定位置
                    rsNurse.Filter = "记录法=1 And 应用方式=1" '标尺数据 心率\脉搏\体温
                    rsNurse.AbsolutePosition = lngPosition
                End If
            End If
            .ItemTag = "K_" & rsNurse!项目序号
        End With
        colMeaSure.Add udtDataItem, "K_" & rsNurse!项目序号
        rsNurse.MoveNext
    Next
    '-------------------------------------------------------------------------------------------------------
    Set colFoot = New Collection
    
    Set colItem = New Collection
    rsNurse.Filter = "记录法=2 And 应用方式=1 And 项目序号 <> 4 " '3-呼吸\血压(4-收缩压/5-舒张压)\7-入液量\9-出液量\10-大便次数\自定义项目..
    For i = 1 To rsNurse.RecordCount
        FuncClearUDT udtDataItem
        strKey = "K_" & rsNurse!项目序号
        strTitle = rsNurse!记录名 & IIf(Nvl(rsNurse!单位) <> "", "(" & rsNurse!单位 & ")", "")
        With udtDataItem
            If rsNurse!项目序号 = 3 Then
                .GraphType = M_CON_TYPE_TICK    '时刻数据
            ElseIf Val(rsNurse!记录频次 & "") > 0 Then
                .GraphType = M_CON_TYPE_DIVISION
                .SplitString = ","
                .SplitCount = Val(rsNurse!记录频次 & "")
                If rsNurse!项目序号 = 5 Then
                    strTitle = "血压" & IIf(Nvl(rsNurse!单位) <> "", "(" & rsNurse!单位 & ")", "")
                End If
            End If
            .Title = strTitle
            .TitleFont = "宋体,9pt"
            .TextColor = IIf(Val(rsNurse!记录色 & "") = 0, "Black", FuncColorRGB(Val(rsNurse!记录色 & "")))
            .BackgroundColor = "255,255,255"
            .Font = "宋体,9pt"
            .ItemTag = strKey & "," & rsNurse!项目序号 & "," & rsNurse!项目类型 & "," & rsNurse!父序号    '索引,项目序号,项目类型(0-普遍,1-波动,2-汇总),父序号
        End With
        colItem.Add udtDataItem, strKey
        rsNurse.MoveNext
    Next
    
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_GROUP
        .Title = "护理项目"
        .GroupPosition = "居左"    '居左\居上
        Set .ListData = colItem
    End With
    colFoot.Add udtDataItem, "K_护理项目"
    '--------------------------------------------------------------------------------------------------------------
    '用药跟踪
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_CONTINUOUS
        .Title = "药品临嘱"
        .TitleFont = "宋体,9pt"
        .TextColor = "Black"
        .Font = "宋体,9pt"
        .Effect = "网格"
        .ShowHotspotEffect = True
        .HotspotColor = FuncColorRGB(COLOR_鲜蓝)
        .HotspotFont = "宋体,9pt"
    End With
    colFoot.Add udtDataItem, M_CON_KEY_用药跟踪_临
    
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_CONTINUOUS
        .Title = "药品长嘱"
        .TitleFont = "宋体,9pt"
        .TextColor = "Black"
        .Font = "宋体,9pt"
        .Effect = "网格"
        .ShowHotspotEffect = True
        .HotspotColor = FuncColorRGB(COLOR_鲜蓝)
        .HotspotFont = "宋体,9pt"
    End With
    colFoot.Add udtDataItem, M_CON_KEY_用药跟踪_长
    '检查
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_LAYOUT
        .Title = "检查"
        .TitleFont = "宋体,9pt"
        .TextColor = "Black"
        .BackgroundColor = "255,255,255"
        .Font = "宋体,9pt"
        .ShowHotspotEffect = True
        .HotspotFont = "宋体,9pt,style=Underline"
        .HotspotColor = FuncColorRGB(COLOR_鲜蓝)
    End With
    colFoot.Add udtDataItem, M_CON_KEY_检查
    '检验
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_LAYOUT
        .Title = "检验"
        .TitleFont = "宋体,9pt"
        .TextColor = "Black"
        .BackgroundColor = "255,255,255"
        .Font = "宋体,9pt"
        .ShowHotspotEffect = True
        .HotspotFont = "宋体,9pt,style=Underline"
        .HotspotColor = FuncColorRGB(COLOR_鲜蓝)
    End With
    colFoot.Add udtDataItem, M_CON_KEY_检验
    '其他医嘱:临时、长期
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_LAYOUT
        .Title = "其他临嘱"
        .TitleFont = "宋体,9pt"
        .TextColor = "Black"
        .BackgroundColor = "White"
        .Font = "宋体,9pt"
        .ShowHotspotEffect = True
        .HotspotColor = FuncColorRGB(COLOR_鲜蓝)
        .HotspotFont = "宋体,9pt"
    End With
    colFoot.Add udtDataItem, M_CON_KEY_其他临嘱
    
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_CONTINUOUS
        .Title = "其他长嘱"
        .TitleFont = "宋体,9pt"
        .TextColor = "Black"
        .BackgroundColor = "White"
        .Font = "宋体,9pt"
        .ShowHotspotEffect = True
        .HotspotColor = FuncColorRGB(COLOR_鲜蓝)
        .HotspotFont = "宋体,9pt"
        .Effect = "网格"
    End With
    colFoot.Add udtDataItem, M_CON_KEY_其他长嘱
    
    '手术
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_LAYOUT
        .Title = "手术"
        .TitleFont = "宋体,9pt"
        .TextColor = "Black"
        .BackgroundColor = "White"
        .Font = "宋体,9pt"
        .ShowHotspotEffect = True
        .HotspotColor = FuncColorRGB(COLOR_鲜蓝)
        .HotspotFont = "宋体,9pt"
    End With
    colFoot.Add udtDataItem, M_CON_KEY_手术
    
    '病历文书
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_LAYOUT
        .Title = "病历文书"
        .TitleFont = "宋体,9pt"
        .TextColor = "Black"
        .BackgroundColor = "White"
        .Font = "宋体,9pt"
        .ShowHotspotEffect = True
        .HotspotFont = "宋体,9pt,style=Underline"
        .HotspotColor = FuncColorRGB(COLOR_深橙)
        .ShowHotspotCursor = True
        .GroupPosition = "居左"
    End With
    colFoot.Add udtDataItem, M_CON_KEY_病历
    
    '---------------------------------------------
    With mudtTimeLine
        Set .colHeaderData = colHead
        Set .colMeasureData = colMeaSure
        Set .colFooterData = colFoot
    End With
    '----------------------------------------------------------------------------------------------------------------------------------------

    If mlng病人ID <> 0 Then
        Call FuncLoadItemNurse
        Call FuncLoadDrug
        Call FuncLoadEMR
        Call FuncLoadPACS_Lis
        Call FuncLoadOperation
        Call FuncLoadInHosDay
        Call FuncLoadOperationAfter
        Call FuncLoadAdvice
    End If
    
    strDesignInfo = FuncMakeXMLDesign(mudtDesign)
    LogWrite "住院一览的调试日志", "" & glngModul, "FuncCreateTimeLine", "Design_Create:" & vbCrLf & strDesignInfo
    strData = FuncMakeXMLTimeLine(mudtTimeLine)
    LogWrite "住院一览的调试日志", "" & glngModul, "FuncCreateTimeLine", "Data_Create:" & vbCrLf & strData
    mtimeLineControl.UpdateDesignInfo strDesignInfo
    mtimeLineControl.UpdateData strData
    mtimeLineControl.RefreshAll
    If mblnMeasureArea Then
        mtimeLineControl.ShowMeasureArea
    Else
        mtimeLineControl.HideMeasureArea
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objMenu As CommandBarPopup
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsSub.VisualTheme = xtpThemeOfficeXP
    With Me.cbsSub.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .UseFadedIcons = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With

    Set cbsSub.Icons = zlCommFun.GetPubIcons
    cbsSub.EnableCustomization False
    cbsSub.ActiveMenuBar.Visible = False
    '------------------------------------------------------------------------------------------------------------------
    
    '弹出菜单
    Set objMenu = cbsSub.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "弹出菜单(&K)", 0, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Img_Look, "观片")
        objControl.IconId = conMenu_Img_Look
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份
    Set objBar = cbsSub.Add("工具栏", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap '+ xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False

    With objBar.Controls
    
        Set mobjPopup = .Add(xtpControlPopup, conMenu_Edit_NewItem, "页面")
        mobjPopup.IconId = conMenu_Edit_Modify
        mobjPopup.Style = xtpButtonIconAndCaption
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Forward, "上一页", -1, False)
        objControl.IconId = conMenu_View_Forward
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Backward, "下一页", -1, False)
        objControl.IconId = conMenu_View_Backward
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Show, IIf(mblnMeasureArea, "隐藏护理", "显示护理"), -1, False)
        objControl.IconId = conMenu_Manage_Up
        objControl.Style = xtpButtonIconAndCaption
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Assistant, "十字线", -1, False)
        objControl.IconId = conMenu_PatholMeal_AddRecord
        objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_Process_Zoom, "放大", -1, False)
        objControl.IconId = conMenu_Process_Zoom
        objControl.Style = xtpButtonIconAndCaption
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Navigatebeginning, "", -1, False)
        objControl.Style = xtpButtonIconAndCaption
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Navigateend, "", -1, False)
        objControl.Style = xtpButtonIconAndCaption
    End With
End Sub

Private Sub FuncLoadPages()
'功能:根据实际情况,添加页面按钮
    Dim objControl As CommandBarControl
    Dim datDate As Date
    Dim strSQL As String, strTmpDate As String
    Dim i As Long, lngDay As Long
    Dim rsTmp As ADODB.Recordset
    
    If Not mobjPopup Is Nothing Then mobjPopup.CommandBar.Controls.DeleteAll
    
    datDate = zlDatabase.Currentdate
    If mlng病人ID = 0 Then
        '未选中病人时 , 界面缺省设置
        Set objControl = mobjPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, "第1页:" & Format(datDate, "YYYY-MM-DD") & "～" & Format(datDate + 6, "YYYY-MM-DD"), -1, False)
        objControl.Parameter = 1 & "," & Format(datDate, "YYYY-MM-DD") & "," & Format(datDate + 6, "YYYY-MM-DD")
        mlngPages = 1
    Else
        '获取病人入院日期,出院日期
        strSQL = "Select a.入院日期, a.出院日期 From 病案主页 A Where a.病人id = [1] And a.主页id = [2]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
        mDatIn = CDate(Format(rsTmp!入院日期, "YYYY-MM-DD"))
        mdatOut = CDate(Format(IIf(Nvl(rsTmp!出院日期, 0) = 0, Format(datDate, "YYYY-MM-DD"), Format(rsTmp!出院日期, "YYYY-MM-DD")), "YYYY-MM-DD"))
        lngDay = (mdatOut - mDatIn) + 1
        If lngDay <= 0 Then
            mlngPages = 1
            mlngDay = 7
        Else
            mlngDay = 7
            mlngPages = lngDay \ mlngDay + IIf(lngDay Mod mlngDay > 0, 1, 0)
        End If
        For i = 1 To mlngPages
            mDatBegin = mDatIn + (i - 1) * mlngDay
            If i < mlngPages Then
                mDatEnd = mDatBegin + (mlngDay - 1)
            Else
                lngDay = mdatOut - mDatBegin
                If lngDay < 7 Then
                    mDatEnd = mDatBegin + 6  '
                Else
                    mDatEnd = mdatOut
                End If
            End If
            strTmpDate = Format(mDatBegin, "YYYY-MM-DD") & "～" & Format(mDatEnd, "YYYY-MM-DD")
            Set objControl = mobjPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, "第" & i & "页:" & strTmpDate, -1, False)
            objControl.Parameter = i & "," & Format(mDatBegin, "YYYY-MM-DD") & "," & Format(mDatEnd, "YYYY-MM-DD")
        Next
    End If
    objControl.Execute
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncLoadDrug()
'功能:
    Dim strSQL As String, strType As String, strColor As String
    Dim str期效 As String, strBegin As String, strEnd As String
    Dim strValue As String, strTag As String, str用法 As String
    Dim strName As String
    Dim rsAdvice As ADODB.Recordset, rsDrug As ADODB.Recordset
    Dim colFoot As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim i As Long, j As Long, n As Long, k As Long, lngPos As Long, lng组ID As Long
    Dim lngRow As Long, lngDay As Long, lngTemp As Long, lngGroupNum As Long
    Dim datBegin As Date, datEnd As Date, datCurr As Date, datTemp As Date
    Dim blnGroup As Boolean
    
    On Error GoTo errH
        
    datBegin = CDate(Format(mDatBegin, "YYYY-MM-DD 00:00:00"))
    datEnd = CDate(Format(mDatEnd, "YYYY-MM-DD 23:59:59"))
    datCurr = zlDatabase.Currentdate
    '用药跟踪取1-西药及中成药整组医嘱记录及皮试医嘱；2-中药配方只取中药服法 诊疗类别=E,操作类型 =4
    strSQL = "Select a.Id, a.相关id, a.医嘱期效, a.诊疗类别, b.名称, a.医嘱内容, b.执行分类, b.操作类型, a.开嘱时间, a.开始执行时间, a.执行终止时间," & vbNewLine & _
            "       Decode(a.首次用量, Null, '', a.首次用量 || b.计算单位 || ':') ||" & vbNewLine & _
            "        Decode(a.单次用量, Null, Null, Decode(Sign(1 - a.单次用量), 1, '0' || a.单次用量, a.单次用量) || b.计算单位) As 单量," & vbNewLine & _
            "       Decode(a.总给予量, Null, Null," & vbNewLine & _
            "               Decode(a.诊疗类别, 'E', Decode(b.操作类型, '4', a.总给予量 || '付', a.总给予量 || b.计算单位), '5'," & vbNewLine & _
            "                       Round(a.总给予量 / d.住院包装, 5) || d.住院单位, '6', Round(a.总给予量 / d.住院包装, 5) || d.住院单位, a.总给予量 || b.计算单位)) As 总量," & vbNewLine & _
            "       a.执行频次, Decode(A.诊疗类别,'E',Decode(Instr('2468',Nvl(B.操作类型,'0')),0,NULL,B.名称),NULL) as 用法 " & vbNewLine & _
            "From 病人医嘱记录 A, 诊疗项目目录 B, 药品规格 D" & vbNewLine & _
            "Where a.诊疗项目id = b.Id(+) And a.收费细目id = d.药品id(+) And a.病人id = [1] And a.主页id = [2] And Nvl(a.婴儿, 0) = [3] And" & vbNewLine & _
            "      (Instr(',5,6,', a.诊疗类别) > 0 Or (a.诊疗类别 = 'E' And b.操作类型 In ('1', '2', '4'))) And" & vbNewLine & _
            "      (a.开始执行时间 Between [4] And [5] OR (a.医嘱期效=0 And a.开始执行时间 < [4] And NVL(a.执行终止时间,[5])>[4]))  And" & vbNewLine & _
            "      ((a.医嘱期效 = 1 And a.医嘱状态 In (3, 8)) Or (a.医嘱期效 = 0 And a.医嘱状态 In (3, 5, 6, 7, 8, 9)))" & vbNewLine & _
            "Order By a.医嘱期效, a.开始执行时间, a.序号"

    Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, mintBaby, datBegin, datEnd)
    For n = 1 To 2
        If n = 1 Then
            rsAdvice.Filter = "医嘱期效=0"
            udtDataItem = mudtTimeLine.colFooterData(M_CON_KEY_用药跟踪_长)
        ElseIf n = 2 Then
            rsAdvice.Filter = "医嘱期效=1"
            udtDataItem = mudtTimeLine.colFooterData(M_CON_KEY_用药跟踪_临)
        End If
        Set colFoot = New Collection
        lngRow = 0: Set rsDrug = InitRS()
        For i = 1 To rsAdvice.RecordCount
            lngPos = rsAdvice.AbsolutePosition '标记当前行
            blnGroup = False
            If (rsAdvice!操作类型 & "" = "4" And rsAdvice!诊疗类别 & "" = "E") Or InStr(",5,6,", rsAdvice!诊疗类别) > 0 Then
                If InStr(",5,6,", rsAdvice!诊疗类别) > 0 Then
                    If lng组ID <> CLng(rsAdvice!相关ID & "") Then
                        blnGroup = True
                        lng组ID = CLng(rsAdvice!相关ID & "")
                        rsAdvice.MoveNext
                        For j = lngPos + 1 To rsAdvice.RecordCount
                            If lng组ID = CLng(rsAdvice!ID & "") Then
                                '0-其他治疗类别,1-输液类,2-注射类,3-皮试,4-口服
                                Select Case Nvl(rsAdvice!执行分类)
                                Case "1"
                                    strType = "[滴]"      'FFC0C0 =BGR   同门诊一览保持一致
                                    strColor = FuncColorRGB(COLOR_浅蓝)
                                Case "2"
                                    strType = "[针]"       '&HFFFFC0
                                    strColor = FuncColorRGB(COLOR_浅青)
                                Case "3"
                                    strType = "[皮]"      '&HC0C0FF
                                    strColor = FuncColorRGB(COLOR_浅红)
                                Case "4"
                                    strType = "[口]"        '&HC0FFC0
                                    strColor = FuncColorRGB(COLOR_浅绿)
                                Case Else
                                    strType = ""
                                    strColor = FuncColorRGB(COLOR_浅橙)           '&HC0E0FF
                                End Select
                                str用法 = rsAdvice!用法 & ""
                                lngGroupNum = j - lngPos
                                Exit For
                            End If
                            rsAdvice.MoveNext
                        Next
                    End If
                    rsAdvice.AbsolutePosition = lngPos
                    strTag = "药嘱内容:" & rsAdvice!医嘱内容 & IIf(Nvl(rsAdvice!总量) = "", "", ",共" & rsAdvice!总量) & IIf(Nvl(rsAdvice!单量) = "", "", ",每次" & rsAdvice!单量) & "," & str用法 & "," & rsAdvice!执行频次
                            
                    strName = rsAdvice!名称 & ""
                ElseIf (rsAdvice!操作类型 & "" = "4" And rsAdvice!诊疗类别 & "" = "E") Then
                    strType = "配"
                    strTag = "药嘱内容:" & rsAdvice!医嘱内容
                    strName = rsAdvice!医嘱内容 & ""
                    strColor = FuncColorRGB(COLOR_浅紫)
                End If
                
                strTag = strTag & vbCrLf & "生效时间:" & Format(rsAdvice!开始执行时间 & "", "YYYY-MM-DD HH:MM:SS")
                
                If Val(rsAdvice!医嘱期效 & "") = 1 Then
                    strEnd = ""
                    strValue = "[临]" & strType & strName
                    strColor = ""
                    strBegin = Format(rsAdvice!开始执行时间, "YYYY-MM-DD ") & Format("2:30:00", "HH:MM:SS")
                    '设置临时药嘱显示列占两列,超过最后一列时,默认终止时间为最后一列的时间
                    If rsAdvice!执行终止时间 & "" = "" Then
                        datTemp = DateAdd("D", 1, CDate(rsAdvice!开始执行时间 & ""))
                    Else
                        datTemp = CDate(rsAdvice!执行终止时间 & "")
                        lngDay = DateDiff("D", CDate(rsAdvice!开始执行时间 & ""), CDate(rsAdvice!执行终止时间 & ""))
                        If lngDay < 1 Then
                            datTemp = CDate(rsAdvice!执行终止时间 & "") + 1
                        End If
                    End If
                    If datTemp > datEnd Then
                        strEnd = Format(datEnd, "YYYY-MM-DD 23:59:59")
                    Else
                        strEnd = Format(datTemp, "YYYY-MM-DD 23:59:59")
                    End If
                Else
                    If rsAdvice!执行终止时间 & "" <> "" Then
                        strTag = strTag & vbCrLf & "终止时间:" & Format(rsAdvice!执行终止时间 & "", "YYYY-MM-DD HH:MM:SS")
                    End If
                    If Between(CDate(Format(rsAdvice!开始执行时间 & "", "YYYY-MM-DD")), datBegin, datEnd) Then
                        strBegin = Format(rsAdvice!开始执行时间, "YYYY-MM-DD ") & Format("2:30:00", "HH:MM:SS")
                    Else
                        strBegin = Format(datBegin, "YYYY-MM-DD ") & Format("2:30:00", "HH:MM:SS")
                    End If
                    
                    If Nvl(rsAdvice!执行终止时间) = "" Then
                        If DateDiff("D", datEnd, datCurr) >= 0 Then
                            datTemp = datEnd
                        Else
                            datTemp = datCurr
                        End If
                    Else
                        If CDate(rsAdvice!执行终止时间) > datEnd Then
                            datTemp = datEnd
                        Else
                            datTemp = CDate(rsAdvice!执行终止时间)
                        End If
                    End If
                    strEnd = Format(datTemp, "YYYY-MM-DD 23:59:59")
                    strValue = "[长]" & strType & rsAdvice!名称
                End If
                
                '用药跟踪空白填补
                If blnGroup Then
                    lngTemp = 0
                    For j = 1 To lngRow
                        rsDrug.Filter = "行号=" & j & " And 日期 >= '" & Format(strBegin, "YYYY-MM-DD") & "'"
                        If rsDrug.RecordCount = 0 Then
                            lngTemp = j   '找到一行空白行
                            If lngGroupNum <= 1 Then
                                Exit For
                            Else
                                '一并给药判断预留空白是否足够
                                For k = 2 To lngGroupNum
                                    rsDrug.Filter = "行号=" & (j + k - 1) & " And 日期 >= '" & Format(strBegin, "YYYY-MM-DD") & "'"
                                    If rsDrug.RecordCount > 0 Then Exit For
                                Next
                                If k <= lngGroupNum Then
                                    j = lngRow + 1: lngTemp = lngRow
                                End If
                                Exit For
                            End If
                        End If
                    Next
                Else
                    If lngTemp < lngRow And lngTemp > 0 Then lngTemp = lngTemp + 1
                End If
                If lngTemp = lngRow Or lngTemp = 0 Then lngRow = lngRow + 1: lngTemp = lngRow
                 
                If Val(rsAdvice!医嘱期效 & "") = 1 Then
                    lngDay = 1
                Else
                    lngDay = DateDiff("D", Format(strBegin, "YYYY-MM-DD"), Format(strEnd, "YYYY-MM-DD"))
                    If lngDay < 1 Then lngDay = 1
                End If
                rsDrug.AddNew
                For j = 0 To lngDay
                    rsDrug!行号 = lngTemp
                    rsDrug!日期 = Format(DateAdd("D", j, Format(strBegin, "YYYY-MM-DD")), "YYYY-MM-DD")
                Next
                rsDrug.UpdateBatch
                
                With udtDataInfo
                    .Value = strValue
                    .RowNumber = lngTemp
                    .RowIndex = rsAdvice!ID
                    .Time = Format(strBegin, "YYYY-MM-DDTHH:MM:SS")
                    .TimeEnd = Format(strEnd, "YYYY-MM-DDTHH:MM:SS")
                    .BackgroundColor = strColor
                    .Group = rsAdvice!相关ID & "" '附加信息记录相关ID
                    .Tip = strTag
                End With
                colFoot.Add udtDataInfo, "_" & (colFoot.Count + 1)
            End If
            rsAdvice.MoveNext
        Next
        If n = 1 Then
            Set udtDataItem.ListData = colFoot
            mudtTimeLine.colFooterData.Remove (M_CON_KEY_用药跟踪_长)
            mudtTimeLine.colFooterData.Add udtDataItem, M_CON_KEY_用药跟踪_长, M_CON_KEY_检查
        Else
            Set udtDataItem.ListData = colFoot
            mudtTimeLine.colFooterData.Remove (M_CON_KEY_用药跟踪_临)
            mudtTimeLine.colFooterData.Add udtDataItem, M_CON_KEY_用药跟踪_临, M_CON_KEY_用药跟踪_长
        End If
    Next

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncLoadInHosDay()
'功能:加载住院天数 普通文本
'算法:从入院日期到当前日期
    Dim colHead As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim lngDay As Long
    Dim i As Long
    
    udtDataItem = mudtTimeLine.colHeaderData(M_CON_KEY_住院天数)
    Set colHead = New Collection
    lngDay = (mdatOut - mDatIn)
    For i = 0 To lngDay
        With udtDataInfo
            .Value = mDatBegin - mDatIn + i
            .Time = Format(mDatBegin + i, "YYYY-MM-DDTHH:MM:SS")
            .BackgroundColor = "White"
            .TextColor = "Black"
        End With
        colHead.Add udtDataInfo, "_" & (colHead.Count + 1)
        If Format(mdatOut, "YYYY-MM-DD") = Format(mDatBegin + i, "YYYY-MM-DD") Then Exit For
    Next
    Set udtDataItem.ListData = colHead
    mudtTimeLine.colHeaderData.Remove (M_CON_KEY_住院天数)
    mudtTimeLine.colHeaderData.Add udtDataItem, M_CON_KEY_住院天数, M_CON_KEY_手术后天数
    
End Sub

Private Sub FuncLoadPACS_Lis()
'功能:检查检验
    Dim colFoot As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strType As String
    Dim strSQL As String
    Dim datBegin As Date, datEnd As Date
    
    datBegin = CDate(Format(mDatBegin, "YYYY-MM-DD 00:00:00"))
    datEnd = CDate(Format(mDatEnd, "YYYY-MM-DD 23:59:59"))
    
    strSQL = "Select a.Id, a.医嘱内容, a.诊疗类别, b.操作类型, a.开始执行时间, Max(c.病历id) As 报告id, Max(c.检查报告id) As 检查报告id," & vbNewLine & _
            "       Decode(Max(Nvl(c.查阅状态, 0)), Min(Nvl(c.查阅状态, 0)), Max(Nvl(c.查阅状态, 0)), 2) As 查阅状态" & vbNewLine & _
            "From 病人医嘱记录 A, 诊疗项目目录 B, 病人医嘱报告 C" & vbNewLine & _
            "Where a.诊疗项目id = b.Id And a.Id = c.医嘱id(+) And a.病人id = [1] And a.主页id = [2] And NVL(a.婴儿,0) =[3] And" & vbNewLine & _
            "      ((a.医嘱期效 = 1 And a.医嘱状态 In (3, 8)) Or (a.医嘱期效 = 0 And a.医嘱状态 In (3, 5, 6, 7, 8, 9))) And" & vbNewLine & _
            "      (a.诊疗类别 = 'D' And a.相关id Is Null Or a.诊疗类别 = 'E' And b.操作类型 = '6')" & vbNewLine & _
            "      And a.开始执行时间 Between [4] And [5] " & vbNewLine & _
            "Group By a.Id, a.医嘱内容, a.序号, a.诊疗类别, b.操作类型, a.开始执行时间 " & vbNewLine & _
            "Order By a.序号"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, mintBaby, datBegin, datEnd)
    
    udtDataItem = mudtTimeLine.colFooterData(M_CON_KEY_检查)
    Set colFoot = New Collection
    rsTmp.Filter = "诊疗类别 = 'D'"
    For i = 1 To rsTmp.RecordCount
        Call FuncClearUDT(udtDataInfo)
        With udtDataInfo
            .RowIndex = rsTmp!ID
            .Value = rsTmp!医嘱内容
            .Time = Format(rsTmp!开始执行时间, "YYYY-MM-DDTHH:MM:SS")
            .BackgroundColor = "White"
            '已出报告 淡蓝色字体显示并加别针图标
            If Not (Val(rsTmp!报告ID & "") = 0 And Val(rsTmp!检查报告ID & "") = 0) Then
                .TextColor = FuncColorRGB(COLOR_鲜蓝)
                .ShowHotspotCursor = True
            Else
                .TextColor = "Black"
                .HotspotFont = "宋体,9pt"
                .ShowHotspotCursor = False
            End If
            .Tag = "D," & rsTmp!报告ID & "," & rsTmp!检查报告ID
        End With
        colFoot.Add udtDataInfo, "_" & (colFoot.Count + 1)
        rsTmp.MoveNext
    Next
    Set udtDataItem.ListData = colFoot
    mudtTimeLine.colFooterData.Remove (M_CON_KEY_检查)
    mudtTimeLine.colFooterData.Add udtDataItem, M_CON_KEY_检查, M_CON_KEY_检验
    '检验
    udtDataItem = mudtTimeLine.colFooterData(M_CON_KEY_检验)
    Set colFoot = New Collection
    rsTmp.Filter = "诊疗类别 = 'E'"
    For i = 1 To rsTmp.RecordCount
        With udtDataInfo
            .RowIndex = rsTmp!ID
            .Value = rsTmp!医嘱内容
            .Time = Format(rsTmp!开始执行时间, "YYYY-MM-DDTHH:MM:SS")
            .BackgroundColor = "White"
            .Tag = "E," & rsTmp!报告ID & "," & rsTmp!检查报告ID
            '已出报告 淡蓝色字体显示并加别针图标
            If Val(rsTmp!报告ID & "") <> 0 Then
                .TextColor = FuncColorRGB(COLOR_鲜蓝)
                .ShowHotspotCursor = True
            Else
                .TextColor = "Black"
                .HotspotFont = "宋体,9pt"  '取消下划线
                .ShowHotspotCursor = False
            End If
        End With
        colFoot.Add udtDataInfo, "_" & (colFoot.Count + 1)
        rsTmp.MoveNext
    Next
    Set udtDataItem.ListData = colFoot
    mudtTimeLine.colFooterData.Remove (M_CON_KEY_检验)
    mudtTimeLine.colFooterData.Add udtDataItem, M_CON_KEY_检验, M_CON_KEY_其他临嘱
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Sub

Private Sub FuncLoadEMR()
'功能:加载病历文件列表
    Dim colFoot As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim strSQL As String, strMsg As String
    Dim rsTmp As ADODB.Recordset
    Dim rsEmr As ADODB.Recordset
    Dim i As Long
    Dim datBegin As Date, datEnd As Date
    
    datBegin = CDate(Format(mDatBegin, "YYYY-MM-DD 00:00:00"))
    datEnd = CDate(Format(mDatEnd, "YYYY-MM-DD 23:59:59"))
     
    '2-住院病历,5-诊断文书,6-知情文件
    strSQL = "Select ID, 病历种类, 病历名称,创建时间, Decode(Nvl(签名级别, 0), 0, '保存(未完成)', 1, '完成', '审签') as 状态 " & vbNewLine & _
            "From 电子病历记录 " & vbNewLine & _
            "Where 病人来源 = 2 And 病历种类 In (2, 5, 6) And 病人id = [1] And 主页id = [2] And NVL(婴儿,0) = [3] " & vbNewLine & _
            "      And 创建时间 Between [4] And [5] " & vbNewLine & _
            "Order By 病历种类, 序号"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, mintBaby, datBegin, datEnd)
    
    udtDataItem = mudtTimeLine.colFooterData(M_CON_KEY_病历)
    Set colFoot = New Collection
    
    For i = 1 To rsTmp.RecordCount
        With udtDataInfo
            .Value = rsTmp!病历名称 & "【" & rsTmp!状态 & "】"
            .Time = Format(rsTmp!创建时间 & "", "YYYY-MM-DDTHH:MM:SS")
            .BackgroundColor = "White"
            .Tag = rsTmp!ID
        End With
        colFoot.Add udtDataInfo, "_" & (colFoot.Count + 1)
        rsTmp.MoveNext
    Next
    '新版病历
    If Not gobjEmr Is Nothing Then
        '新版病历提供接口：GetInEPRRecord(病人ID,主页ID,RS)返回每次就诊的病历情况（Id, Title,Creat_Time,STATUS( 编辑中、已签名、审订中、已审签)）。
        On Error Resume Next
        strMsg = gobjEmr.GetInEPRRecord(mlng病人ID, mlng主页ID, rsEmr)
        err.Clear: On Error GoTo 0
        If Not rsEmr Is Nothing Then
            For i = 1 To rsEmr.RecordCount
                With udtDataInfo
                    .Value = rsEmr!Title
                    .Time = Format(rsEmr!Creat_Time & "", "YYYY-MM-DDTHH:MM:SS")
                    If rsEmr.Fields.Count = 4 Then
                        If UCase(rsEmr.Fields(3).Name) = UCase("Status") Then
                            .Value = rsEmr!Title & IIf(rsEmr!Status & "" <> "", "【" & rsEmr!Status & "】", "")
                        End If
                    End If
                    .BackgroundColor = "White"
                    .Tag = rsEmr!ID   '新版病历ID值
                End With
                colFoot.Add udtDataInfo, "_" & (colFoot.Count + 1)
                rsEmr.MoveNext
            Next
        End If
    End If
    
    Set udtDataItem.ListData = colFoot
    mudtTimeLine.colFooterData.Remove (M_CON_KEY_病历)
    mudtTimeLine.colFooterData.Add udtDataItem, M_CON_KEY_病历
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
        
End Sub

Private Sub FuncLoadOperation()
'功能:加载手术医嘱
    Dim colFoot As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim datBegin As Date, datEnd As Date
    
    datBegin = CDate(Format(mDatBegin, "YYYY-MM-DD 00:00:00"))
    datEnd = CDate(Format(mDatEnd, "YYYY-MM-DD 23:59:59"))
    
    '手术组医嘱
    strSQL = "Select a.Id, a.手术时间, a.医嘱内容, b.名称 " & vbNewLine & _
            "From 病人医嘱记录 A,诊疗项目目录 B " & vbNewLine & _
            "Where a.诊疗项目id = b.Id And a.病人id = [1] And a.主页id = [2] And NVL(a.婴儿,0) =[3] And a.诊疗类别 = 'F' And Nvl(a.相关id, 0) = 0 And a.医嘱状态 In (3, 8) " & vbNewLine & _
            " And 手术时间 Between [4] and [5] "

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, mintBaby, datBegin, datEnd)
    
    udtDataItem = mudtTimeLine.colFooterData(M_CON_KEY_手术)
    Set colFoot = New Collection
    
    For i = 1 To rsTmp.RecordCount
        With udtDataInfo
            .Value = rsTmp!名称 & ""
            .Time = Format(rsTmp!手术时间 & "", "YYYY-MM-DDTHH:MM:SS")
            .RowIndex = rsTmp!ID
            .Tag = rsTmp!医嘱内容 & ""
        End With
        colFoot.Add udtDataInfo, "_" & (colFoot.Count + 1)
        rsTmp.MoveNext
    Next
    
    Set udtDataItem.ListData = colFoot
    mudtTimeLine.colFooterData.Remove (M_CON_KEY_手术)
    mudtTimeLine.colFooterData.Add udtDataItem, M_CON_KEY_手术, M_CON_KEY_病历
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncLoadOperationAfter()
'功能:加载手术天数
    Dim colHead As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim datBegin As Date, datEnd As Date
    '由于分页限制,需要将开始时间设置成入院时间和出院时间，否则当分页时间段没有手术记录时,将不会显示手术后天数
    datBegin = CDate(Format(mDatIn, "YYYY-MM-DD 00:00:00"))
    datEnd = CDate(Format(mdatOut, "YYYY-MM-DD 23:59:59"))
    
    '手术组医嘱
    strSQL = "Select a.手术时间 " & vbNewLine & _
            "From 病人医嘱记录 A " & vbNewLine & _
            "Where a.病人id = [1] And a.主页id = [2] And NVL(a.婴儿,0) =[3] And a.诊疗类别 = 'F' And Nvl(a.相关id, 0) = 0 And a.医嘱状态 In (3, 8) " & vbNewLine & _
            " And 手术时间 Between [4] and [5] "

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, mintBaby, datBegin, datEnd)
    
    udtDataItem = mudtTimeLine.colHeaderData(M_CON_KEY_手术后天数)
    Set colHead = New Collection
    For i = 1 To rsTmp.RecordCount
        With udtDataInfo
            .Time = Format(rsTmp!手术时间 & "", "YYYY-MM-DDTHH:MM:SS")
            .TimeEnd = Format(mdatOut, "YYYY-MM-DDTHH:MM:SS")      '
        End With
        colHead.Add udtDataInfo, "_" & (colHead.Count + 1)
        rsTmp.MoveNext
    Next
    
    Set udtDataItem.ListData = colHead
    mudtTimeLine.colHeaderData.Remove (M_CON_KEY_手术后天数)
    mudtTimeLine.colHeaderData.Add udtDataItem, M_CON_KEY_手术后天数, , M_CON_KEY_住院天数
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncLoadAdvice()
'功能:加载其他医嘱
    Dim colFoot As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim strSQL As String, strBegin As String, strEnd As String
    Dim rsTmp As ADODB.Recordset
    Dim rsDrug As ADODB.Recordset
    Dim i As Long, j As Long, lngTemp As Long
    Dim lngRow As Long, lngDay As Long
    Dim lngColor As Long
    Dim datBegin As Date, datEnd As Date, datCurr As Date, datTemp As Date
    
    '加载其他医嘱:临时医嘱和长期医嘱分别加载
    datBegin = CDate(Format(mDatBegin, "YYYY-MM-DD 00:00:00"))
    datEnd = CDate(Format(mDatEnd, "YYYY-MM-DD 23:59:59"))
    datCurr = zlDatabase.Currentdate
    strSQL = "Select a.ID,a.医嘱内容, b.名称,a.开始执行时间,a.执行终止时间,a.医嘱期效,诊疗类别 " & vbNewLine & _
        "From 病人医嘱记录 A, 诊疗项目目录 B" & vbNewLine & _
        "Where a.诊疗项目id = b.Id And a.病人id = [1] And a.主页id = [2] And NVL(a.婴儿,0) =[3]  And Not a.诊疗类别 In ('G', 'F', 'D', 'C', '5', '6', '7') And" & vbNewLine & _
        "      Not (Nvl(b.操作类型, 0) In ('2', '3', '4', '6', '8') And a.诊疗类别 = 'E') And Nvl(相关id, 0) = 0 And" & vbNewLine & _
        "      (a.开始执行时间 Between [4] And [5] OR (a.医嘱期效=0 And a.开始执行时间 < [4] And NVL(a.执行终止时间,[5])>[4]))  And" & vbNewLine & _
        "      ((a.医嘱期效 = 1 And a.医嘱状态 In (3, 8)) Or (a.医嘱期效 = 0 And a.医嘱状态 In (3, 5, 6, 7, 8, 9)))" & vbNewLine & _
        "Order By a.开始执行时间, a.序号"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, mintBaby, datBegin, datEnd)
    
    rsTmp.Filter = "医嘱期效 = 1"
    udtDataItem = mudtTimeLine.colFooterData(M_CON_KEY_其他临嘱)
    Set colFoot = New Collection
    
    For i = 1 To rsTmp.RecordCount
        Select Case UCase(rsTmp!诊疗类别 & "")
        Case "H"
            lngColor = COLOR_浅绿
        Case "L"
            lngColor = COLOR_浅黄
        Case "Z"
            lngColor = COLOR_淡红
        Case "I"  '膳食
            lngColor = COLOR_浅蓝
        Case "K"   '输血
            lngColor = COLOR_浅红
        Case "M", "4"   '材料,卫材
            lngColor = COLOR_浅青
        Case "E"
            lngColor = COLOR_淡蓝
        Case Else
            lngColor = vbWhite
        End Select
        With udtDataInfo
            .Value = rsTmp!名称 & ""
            .Time = Format(rsTmp!开始执行时间 & "", "YYYY-MM-DDTHH:MM:SS")
            .BackgroundColor = FuncColorRGB(lngColor)
            .RowIndex = rsTmp!ID
            .Tip = "医嘱内容:" & rsTmp!名称 & "" & vbCrLf & _
                   "生效时间:" & Format(rsTmp!开始执行时间 & "", "YYYY-MM-DDTHH:MM:SS")
        End With
        colFoot.Add udtDataInfo, "_" & (colFoot.Count + 1)
        rsTmp.MoveNext
    Next
    
    Set udtDataItem.ListData = colFoot
    mudtTimeLine.colFooterData.Remove (M_CON_KEY_其他临嘱)
    mudtTimeLine.colFooterData.Add udtDataItem, M_CON_KEY_其他临嘱, M_CON_KEY_其他长嘱
    
    rsTmp.Filter = "医嘱期效 = 0"
    udtDataItem = mudtTimeLine.colFooterData(M_CON_KEY_其他长嘱)
    Set colFoot = New Collection
    Set rsDrug = InitRS
    For i = 1 To rsTmp.RecordCount
        If Between(CDate(rsTmp!开始执行时间 & ""), datBegin, datEnd) Then
            strBegin = Format(rsTmp!开始执行时间, "YYYY-MM-DD ")
        Else
            strBegin = Format(datBegin, "YYYY-MM-DD ")
        End If
        If Nvl(rsTmp!执行终止时间) = "" Then
            If DateDiff("D", datEnd, datCurr) >= 0 Then
                datTemp = datEnd
            Else
                datTemp = datCurr
            End If
        Else
            If CDate(rsTmp!执行终止时间) > datEnd Then
                datTemp = datEnd
            Else
                datTemp = CDate(rsTmp!执行终止时间)
            End If
        End If
        strEnd = Format(datTemp, "YYYY-MM-DD 23:59:59")
        Select Case UCase(rsTmp!诊疗类别 & "")
        Case "H"
            lngColor = COLOR_浅绿
        Case "L"
            lngColor = COLOR_浅黄
        Case "Z"
            lngColor = COLOR_淡红
        Case "I"  '膳食
            lngColor = COLOR_浅蓝
        Case "K"   '输血
            lngColor = COLOR_浅红
            lngColor = &HFF
        Case "M"    '卫材
            lngColor = COLOR_浅青
        Case "E"
            lngColor = COLOR_淡蓝
        Case Else
            lngColor = vbWhite
        End Select
        lngTemp = 0
        For j = 1 To lngRow
            rsDrug.Filter = "行号=" & j & " And 日期 >= '" & Format(strBegin, "YYYY-MM-DD") & "'"
            If rsDrug.RecordCount = 0 Then
                lngTemp = j
                Exit For
            End If
        Next
        If j > lngRow Then lngRow = lngRow + 1: lngTemp = lngRow
        If strEnd = "" Then
            lngDay = 1
        Else
            lngDay = DateDiff("D", Format(strBegin, "YYYY-MM-DD"), Format(strEnd, "YYYY-MM-DD"))
            If lngDay < 1 Then lngDay = 1
        End If
        rsDrug.AddNew
        For j = 0 To lngDay
            rsDrug!行号 = lngTemp
            rsDrug!日期 = Format(DateAdd("D", j, Format(strBegin, "YYYY-MM-DD")), "YYYY-MM-DD")
        Next
        rsDrug.UpdateBatch
    
        With udtDataInfo
            .RowNumber = lngTemp
            .Value = rsTmp!名称 & ""
            .Time = Format(strBegin, "YYYY-MM-DDTHH:MM:SS")
            .TimeEnd = Format(strEnd, "YYYY-MM-DDTHH:MM:SS")
            .BackgroundColor = FuncColorRGB(lngColor)
            .Tip = "医嘱内容:" & rsTmp!医嘱内容 & vbCrLf & _
                   "生效时间:" & Format(rsTmp!开始执行时间 & "", "YYYY-MM-DD HH:MM:SS") & IIf(rsTmp!执行终止时间 & "" <> "", vbCrLf & "终止时间:" & Format(rsTmp!执行终止时间 & "", "YYYY-MM-DD HH:MM:SS"), "")
        End With
        colFoot.Add udtDataInfo, "_" & (colFoot.Count + 1)
        rsTmp.MoveNext
    Next
    
    Set udtDataItem.ListData = colFoot
    mudtTimeLine.colFooterData.Remove (M_CON_KEY_其他长嘱)
    mudtTimeLine.colFooterData.Add udtDataItem, M_CON_KEY_其他长嘱, M_CON_KEY_手术
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncLoadItemNurse()
'功能:加载护理项目
    Dim colTmp As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim strSQL As String, strKey  As String, strTemp As String
    Dim strRecordID As String
    Dim rsTmp As ADODB.Recordset, rsCopy As ADODB.Recordset
    Dim arrTag As Variant
    Dim i As Long, j As Long
    Dim lngRow As Long
    Dim datBegin As Date, datEnd As Date

    datBegin = CDate(Format(mDatBegin, "YYYY-MM-DD 00:00:00"))
    datEnd = CDate(Format(mDatEnd, "YYYY-MM-DD 23:59:59"))
    
    '加载护理项目
    strSQL = "Select c.发生时间,d.记录ID, d.项目序号, d.记录内容, d.记录标记,d.体温部位, f.记录法, f.记录符,NVL(f.记录频次,2) as 记录频次 " & vbNewLine & _
            "From 病人护理文件 A, 病历文件列表 B, 病人护理数据 C, 病人护理明细 D, 体温记录项目 F" & vbNewLine & _
            "Where a.格式id = b.Id And a.Id = c.文件id And c.Id = d.记录id And d.项目序号 = f.项目序号 And a.病人id = [1] And a.主页id = [2] And" & vbNewLine & _
            "      NVL(a.婴儿,0) = [3] And b.种类 = 3 And b.保留 = -1 And c.发生时间 Between [4] And [5]" & vbNewLine & _
            "Order By f.记录法, d.项目序号, c.发生时间,d.记录标记 "
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, mintBaby, datBegin, datEnd)
    '心率,体温,脉搏
    lngRow = mudtTimeLine.colMeasureData.Count
    For i = 1 To lngRow
        udtDataItem = mudtTimeLine.colMeasureData(i)
        strKey = udtDataItem.ItemTag
        Set colTmp = New Collection
        If strKey = "K_1" Then
            '体温：37/38  升降温表示法,同一时间点（同一记录ID）,产生两行数据
            rsTmp.Filter = "项目序号=1"
            strRecordID = ""
            Call FuncClearUDT(udtDataInfo)
            For j = 1 To rsTmp.RecordCount
                If j = 1 Then strRecordID = rsTmp!记录ID
                If strRecordID <> rsTmp!记录ID Then
                    strRecordID = rsTmp!记录ID
                    colTmp.Add udtDataInfo, "_" & (colTmp.Count + 1)
                    Call FuncClearUDT(udtDataInfo)
                End If
                With udtDataInfo
                    If rsTmp!记录标记 = 0 Then
                        .Time = Format(rsTmp!发生时间 & "", "YYYY-MM-DDTHH:MM:SS")
                        .NumberValue = IIf(Val(rsTmp!记录内容 & "") = 0, "", Val(rsTmp!记录内容 & ""))
                        arrTag = Split(rsTmp!记录符 & "", ",")
                        If rsTmp!体温部位 = "口温" Then
                            .LegendType = arrTag(0)
                        ElseIf rsTmp!体温部位 = "腋温" Then
                            .LegendType = arrTag(1)
                        ElseIf rsTmp!体温部位 = "肛温" Then
                            .LegendType = arrTag(2)
                        ElseIf rsTmp!体温部位 = "耳温" Then
                            .LegendType = arrTag(3)
                        ElseIf rsTmp!体温部位 = "额温" Then
                            .LegendType = arrTag(4)
                        End If
                    ElseIf rsTmp!记录标记 = 1 Then
                        .BalloonValue = rsTmp!记录内容 & ""
                        .BalloonLegendType = "空心圆"
                    End If
                End With
                If rsTmp.RecordCount = j Then
                    colTmp.Add udtDataInfo, "_" & (colTmp.Count + 1)
                    Call FuncClearUDT(udtDataInfo)
                End If
                rsTmp.MoveNext
            Next
        ElseIf strKey = "K_2" Then
            '脉搏
            rsTmp.Filter = "项目序号 =-1"
            Set rsCopy = zlDatabase.CopyNewRec(rsTmp)
            rsTmp.Filter = "项目序号 = 2"
            rsTmp.Sort = "发生时间"
            Do While Not rsTmp.EOF
                Call FuncClearUDT(udtDataInfo)
                With udtDataInfo
                    .Time = Format(rsTmp!发生时间 & "", "YYYY-MM-DDTHH:MM:SS")
                    .NumberValue = rsTmp!记录内容 & ""
                    arrTag = Split(rsTmp!记录符 & "", ",")
                    '起搏器缺省是“H”,缺省取设置值
                    If rsTmp!体温部位 = "起搏器" Then
                        .LegendType = "H"     'H
                    Else
                        .LegendType = arrTag(0) '
                    End If
                    
                    rsCopy.Filter = "发生时间 =#" & rsTmp!发生时间 & "#"
                    If Not rsCopy.EOF Then
                        .ShadowValue = rsCopy!记录内容 & ""
                        .ShadowLegendType = rsCopy!记录符 & ""
                    Else
                        .ShadowValue = rsTmp!记录内容 & ""
                        .ShadowLegendType = "点"
                    End If
                End With
                colTmp.Add udtDataInfo, "_" & (colTmp.Count + 1)
                rsTmp.MoveNext
            Loop
        ElseIf strKey = "K_-1" Then
            rsTmp.Filter = "项目序号 =-1"
            rsTmp.Sort = "发生时间"
            Call FuncClearUDT(udtDataInfo)
            For j = 1 To rsTmp.RecordCount
                With udtDataInfo
                    .Time = Format(rsTmp!发生时间 & "", "YYYY-MM-DDTHH:MM:SS")
                    .NumberValue = rsTmp!记录内容 & ""
                    .LegendType = rsTmp!记录符 & ""
                End With
                colTmp.Add udtDataInfo, "_" & (colTmp.Count + 1)
                Call FuncClearUDT(udtDataInfo)
                rsTmp.MoveNext
            Next
        End If
        If i < lngRow Then strTemp = mudtTimeLine.colMeasureData(i + 1).ItemTag
        Set udtDataItem.ListData = colTmp
        mudtTimeLine.colMeasureData.Remove (strKey)
        If i < lngRow Then
            mudtTimeLine.colMeasureData.Add udtDataItem, strKey, strTemp
        Else
            mudtTimeLine.colMeasureData.Add udtDataItem, strKey
        End If
    Next
    
    '血压
    Call FuncGetItemNurse(datBegin, datEnd, rsTmp)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function FuncColorRGB(lngColor As Long) As String
'功能:返回RGB颜色值的字符串
    Dim Color(2) As Byte
    Color(0) = (lngColor Mod 256)
    Color(1) = ((lngColor Mod 65536) \ 256)
    Color(2) = (lngColor \ 65536)
    FuncColorRGB = Color(0) & "," & Color(1) & "," & Color(2)
End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
'功能:设置医嘱清单的字体大小
'入参:bytSize：0-小(缺省)，1-大
    Dim strDesignInfo As String
    Dim strData As String
    
    mbytFont = bytSize

    Call zlControl.SetPubFontSize(Me, bytSize)

    strDesignInfo = FuncMakeXMLDesign(mudtDesign)
    LogWrite "住院一览的调试日志", "" & glngModul, "SetFontSize", "Design_FontSize:" & vbCrLf & strDesignInfo
    strData = FuncMakeXMLTimeLine(mudtTimeLine)
    LogWrite "住院一览的调试日志", "" & glngModul, "SetFontSize", "Data_FontSize:" & vbCrLf & strData
    mtimeLineControl.UpdateDesignInfo strDesignInfo
    mtimeLineControl.UpdateData strData
    mtimeLineControl.RefreshAll
    If mblnMeasureArea Then
        mtimeLineControl.ShowMeasureArea
    Else
        mtimeLineControl.HideMeasureArea
    End If
End Sub

Public Sub zlExecuteCommandBars()
'功能:住院医生站调用
    Dim objContrl As CommandBarControl
    Set objContrl = cbsSub.FindControl(xtpControlButton, conMenu_Process_Zoom)
    objContrl.Execute
End Sub

Private Sub mtimeLineControl_DataMouseDoubleClick(ByVal sender As Variant, ByVal e As ZLSoft_BusinessHome_ClientControl_TimeLineBase.IEventArgsDataInfo)
'事件申明用于避免鼠标双击抛异常
End Sub

Private Sub mtimeLineControl_DataTitleMouseDoubleClick(ByVal sender As Variant, ByVal e As ZLSoft_BusinessHome_ClientControl_TimeLineBase.IEventArgsDataInfo)
'事件申明用于避免鼠标双击抛异常
End Sub

Private Sub mtimeLineControl_DateMouseDoubleClick(ByVal sender As Variant, ByVal e As ZLSoft_BusinessHome_ClientControl_TimeLineBase.IEventArgsDateInfo)
'事件申明用于避免鼠标双击抛异常
End Sub

Private Sub mtimeLineControl_TimeLineMouseClick(ByVal sender As Variant, ByVal e As ZLSoft_BusinessHome_ClientControl_TimeLineBase.IEventArgsTimeLineMouse)
'事件申明用于避免鼠标单击抛异常
End Sub

Private Function FuncCompareTime(ByVal strBegin As String, ByVal strEnd As String, ByVal strCheckVal As String) As Long
'功能:返回当前时间与中间时间(起始时间和终止时间的中间时间)相差的秒数。
'参数:
'strBegin-起始时间
'strEnd-终止时间
'strCheckVal-要对比的时间

    Dim lngDiff As Double
    Dim datTemp As Date
    
    lngDiff = DateDiff("s", CDate(strBegin), CDate(strEnd)) \ 2
    datTemp = DateAdd("s", lngDiff, CDate(strBegin))

    FuncCompareTime = Abs(DateDiff("s", datTemp, CDate(strCheckVal)))
End Function


Private Function FuncGetNodeSN(ByVal strSN As String) As String
'功能:返回汇总节点下所有序号
    Dim strRet As String
    Dim lngPos As Long
    
    mrs汇总项目.Filter = "父序号=" & strSN
    
    If mrs汇总项目.RecordCount = 0 Then
        strRet = strSN
    Else
        Do While Not mrs汇总项目.EOF
            lngPos = mrs汇总项目.AbsolutePosition
            strRet = strRet & "," & FuncGetNodeSN(mrs汇总项目!序号 & "")
            mrs汇总项目.Filter = "父序号=" & strSN
            mrs汇总项目.AbsolutePosition = lngPos
            mrs汇总项目.MoveNext
        Loop
        strRet = Mid(strRet, 2)
    End If
    If InStr("," & strRet & ",", "," & strSN & ",") = 0 Then strRet = strSN & "," & strRet
    FuncGetNodeSN = strRet
End Function

Private Function InitRS(Optional ByVal bytFunc As Byte = 1) As ADODB.Recordset
'功能:构造记录集
    Dim rs As ADODB.Recordset
    Dim strFields As String
    Dim strFieldName As String
    Dim lngLen As Long
    Dim FieldType As DataTypeEnum
    Dim i As Long, j As Long
    
    Dim arrField As Variant
    Dim arrSubFeld As Variant '字段名称|字段类型|字段长度 缺省字段类型 为adVarChar
    
    If bytFunc = 1 Then
        strFields = "行号|adBigInt|18,日期|adVarChar|10"
    Else
        strFields = "StartDate|adVarChar|20,TickWidth|adVarChar|20"
    End If
    
    Set rs = New ADODB.Recordset
    '-----------------------------------------
    With rs.Fields
        arrField = Split(strFields, ",")
        For i = LBound(arrField) To UBound(arrField)
            arrSubFeld = Split(arrField(i), "|")
            strFieldName = arrSubFeld(0)
            Select Case UCase(arrSubFeld(1) & "")
            Case UCase("adVarChar")
                FieldType = adVarChar
            Case UCase("adBigInt")
                FieldType = adBigInt
            End Select
            lngLen = Val(arrSubFeld(2))
            .Append strFieldName, FieldType, lngLen
        Next
    End With
    '---------------------------------------
    rs.CursorLocation = adUseClient
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenStatic
    rs.Open
    '----------------------------------
    Set InitRS = rs
End Function

Private Function FuncGetSubString(ByVal strInput As String, ByVal lngWidth As Long, Optional ByVal lngStart As Long = 9) As String
'功能:根据传入的长度截取字符
    Dim strRet As String
    Dim i As Long
    
    If Me.TextWidth(strInput) > lngWidth Then
        strRet = Mid(strInput, 1, lngStart)
        For i = lngStart + 1 To Len(strInput)
            If Me.TextWidth(strRet & Mid(strInput, i, 1) & "...") >= lngWidth Then
                strRet = strRet & "..."
                Exit For
            Else
                strRet = strRet & Mid(strInput, i, 1)
            End If
        Next
    Else
        strRet = strInput
    End If
    FuncGetSubString = strRet
End Function

Private Sub FuncGetItemNurse(ByVal datBegin As Date, ByVal datEnd As Date, ByVal rsItem As ADODB.Recordset)
'功能:
    Dim i As Long, j As Long, k As Long
    Dim lngRow As Long
    Dim lngPos As Long, lngPosOne As Long
    Dim lngMinDiff As Long, lngDateDiff As Long

    Dim bytItem As Byte '0-普通项目;1-波动项目;2-汇总项目,
    Dim bytSplitNum As Byte
    Dim datCurr As Date
    
    Dim colTmp As Collection, colblood As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim arrTmp As Variant, varItem As Variant
    Dim strFreB As String, strFreE As String, strMin As String, strMax As String
    Dim strBegin As String, strEnd As String
    Dim strTemp As String, strItem As String, strPrveTime As String
    Dim rsCopy As ADODB.Recordset
    Dim blnDo As Boolean
    
    On Error GoTo errH
    lngRow = mudtTimeLine.colFooterData("K_护理项目").ListData.Count
    For i = 1 To lngRow
        udtDataItem = mudtTimeLine.colFooterData("K_护理项目").ListData(i)
        Set colTmp = New Collection
        arrTmp = Split(udtDataItem.ItemTag, ",")   '索引,项目序号,波动,汇总,父序号
        bytItem = Val(arrTmp(2))
        
        If arrTmp(0) = "K_3" Then  '呼吸
            rsItem.Filter = "项目序号 =3"
            For j = 1 To rsItem.RecordCount
                With udtDataInfo
                    .Value = rsItem!记录内容 & ""
                    .Time = Format(rsItem!发生时间 & "", "YYYY-MM-DDTHH:MM:SS")
                End With
                colTmp.Add udtDataInfo, "_" & (colTmp.Count + 1)
                rsItem.MoveNext
            Next
        Else
            If arrTmp(1) = "5" Then
                Set colblood = New Collection
                arrTmp(1) = "4"  '收缩压  先将收缩压缓存
            End If
LineBlood:
            If bytItem = 2 Then
                strTemp = FuncGetNodeSN(arrTmp(1))
                varItem = Split(strTemp, ","): strItem = ""
                For j = LBound(varItem) To UBound(varItem)
                    strItem = strItem & " OR 项目序号 =" & varItem(j)
                Next
                rsItem.Filter = Mid(strItem, 4)
            Else
                rsItem.Filter = "项目序号 = " & arrTmp(1)
            End If
            rsItem.Sort = "发生时间,项目序号"
            If rsItem.RecordCount > 0 Then
                Set rsCopy = zlDatabase.CopyNewRec(rsItem)  '缓存该项目本页所有数据
                '频次
                bytSplitNum = Val(rsCopy!记录频次 & "")
                If bytItem = 0 Then
                    mrsFrequency.Filter = "频次=" & bytSplitNum
                    If mrsFrequency.RecordCount > 0 Then
                        mrsFrequency.MoveFirst
                        strBegin = mrsFrequency!开始 & ""
                        mrsFrequency.MoveLast
                        strEnd = mrsFrequency!结束 & ""
                    Else
                        strBegin = "00:00"
                        strFreE = "23:59"
                    End If
                ElseIf bytItem = 1 Or bytItem = 2 Then
                    If bytSplitNum = 1 Then
                        mrs汇总时段.Filter = "类别 = 3"
                    Else
                        mrs汇总时段.Filter = "类别 < 3"
                    End If
                    If mrs汇总时段.RecordCount > 0 Then
                        mrs汇总时段.MoveFirst
                        strBegin = mrs汇总时段!开始 & ""
                        mrs汇总时段.MoveLast
                        strEnd = mrs汇总时段!结束 & ""
                    Else
                        strBegin = "00:00"
                        strEnd = "23:59"
                    End If
                End If
                datCurr = datBegin
                Do While datCurr <= datEnd
                    If strBegin <= strEnd Then
                        rsCopy.Filter = "发生时间 >= #" & Format(datCurr, "YYYY-MM-DD ") & Format(strBegin, "HH:MM:SS") & "# And 发生时间 <= #" & Format(datCurr, "YYYY-MM-DD ") & Format(strEnd, "HH:MM:SS") & "#"
                    Else
                        rsCopy.Filter = "发生时间 >= #" & Format(datCurr, "YYYY-MM-DD ") & Format(strBegin, "HH:MM:SS") & "# And 发生时间 <= #" & Format(DateAdd("d", 1, datCurr), "YYYY-MM-DD ") & Format(strEnd, "HH:MM:SS") & "#"
                    End If
                    
                    If rsCopy.RecordCount > 0 Then
                        rsCopy.Sort = "发生时间,项目序号"
                        strItem = ""
                        For k = 1 To bytSplitNum - 1
                            strItem = strItem & ","
                        Next
                        If strItem = "" Then
                            varItem = Array("")
                        Else
                            varItem = Split(strItem, ",")
                        End If
                    End If
                    
                    strItem = "": strMin = "": strMax = ""
                    Do While Not rsCopy.EOF
                        If bytItem = 0 Then
                            mrsFrequency.Filter = "频次=" & bytSplitNum
                            lngPos = rsCopy.AbsolutePosition
                            For k = 1 To mrsFrequency.RecordCount
                                If mrsFrequency!开始 & "" <= mrsFrequency!结束 & "" Then
                                    strFreB = Format(Format(datCurr, "YYYY-MM-DD ") & mrsFrequency!开始, "YYYY-MM-DD HH:MM:SS")
                                    strFreE = Format(Format(datCurr, "YYYY-MM-DD ") & mrsFrequency!结束, "YYYY-MM-DD HH:MM:SS")
                                Else
                                    strFreB = Format(Format(datCurr, "YYYY-MM-DD ") & mrsFrequency!开始, "YYYY-MM-DD HH:MM:SS")
                                    strFreE = Format(Format(DateAdd("d", 1, datCurr), "YYYY-MM-DD ") & mrsFrequency!结束, "YYYY-MM-DD HH:MM:SS")
                                End If
                                If Between(CDate(rsCopy!发生时间), CDate(strFreB), CDate(strFreE)) Then
                                    If varItem(mrsFrequency!序号 - 1) <> "" Then Exit For
                                    If mrsFrequency!类别 = 1 Then
                                        '取第一条
                                        varItem(mrsFrequency!序号 - 1) = rsCopy!记录内容
                                        Exit For
                                    ElseIf mrsFrequency!类别 = 2 Then
                                        '取中间一条
                                        lngPosOne = lngPos
                                        lngMinDiff = -1
                                        Do While Not rsCopy.EOF
                                            If Between(CDate(rsCopy!发生时间), CDate(strFreB), CDate(strFreE)) Then
                                                lngDateDiff = FuncCompareTime(strFreB, strFreE, Format(rsCopy!发生时间 & "", "YYYY-MM-DD HH:MM:SS"))
                                                If lngMinDiff >= lngDateDiff Or lngMinDiff = -1 Then
                                                    lngMinDiff = lngDateDiff
                                                    lngPosOne = rsCopy.AbsolutePosition
                                                Else
                                                    Exit Do
                                                End If
                                            Else
                                                Exit Do
                                            End If
                                            rsCopy.MoveNext
                                        Loop
                                        rsCopy.AbsolutePosition = lngPosOne
                                        varItem(mrsFrequency!序号 - 1) = rsCopy!记录内容
                                        Exit For
                                    ElseIf mrsFrequency!类别 = 3 Then
                                        '取最后一条
                                        rsCopy.MoveNext
                                        If rsCopy.EOF Then
                                            rsCopy.AbsolutePosition = lngPos
                                            varItem(mrsFrequency!序号 - 1) = rsCopy!记录内容
                                        ElseIf Format(rsCopy!发生时间 & "", "YYYY-MM-DD HH:MM:SS") > strFreE Then
                                            rsCopy.AbsolutePosition = lngPos
                                            varItem(mrsFrequency!序号 - 1) = rsCopy!记录内容
                                        Else
                                            rsCopy.AbsolutePosition = lngPos
                                        End If
                                        Exit For
                                    End If
                                End If
                                mrsFrequency.MoveNext
                            Next
                        ElseIf bytItem = 1 Or bytItem = 2 Then
                            '波动项目、汇总项目
                            If bytSplitNum = 1 Then
                                mrs汇总时段.Filter = "类别 = 3"
                            Else
                                mrs汇总时段.Filter = "类别 < 3"
                            End If
                            If bytItem = 1 Then
                                For k = 1 To mrs汇总时段.RecordCount
                                    If mrs汇总时段!开始 & "" <= mrs汇总时段!结束 & "" Then
                                        strFreB = Format(Format(datCurr, "YYYY-MM-DD ") & mrs汇总时段!开始, "YYYY-MM-DD HH:MM:SS")
                                        strFreE = Format(Format(datCurr, "YYYY-MM-DD ") & mrs汇总时段!结束, "YYYY-MM-DD HH:MM:SS")
                                    Else
                                        strFreB = Format(Format(datCurr, "YYYY-MM-DD ") & mrs汇总时段!开始, "YYYY-MM-DD HH:MM:SS")
                                        strFreE = Format(Format(DateAdd("d", 1, datCurr), "YYYY-MM-DD ") & mrs汇总时段!结束, "YYYY-MM-DD HH:MM:SS")
                                    End If
                                    If varItem(k - 1) = "" Then
                                        If Between(Format(rsCopy!发生时间 & "", "YYYY-MM-DD HH:MM:SS"), strFreB, strFreE) Then
                                            If Val(strMax) < Val(rsCopy!记录内容 & "") Or strMax = "" Then strMax = rsCopy!记录内容 & ""
                                            If Val(strMin) > Val(rsCopy!记录内容 & "") Or strMin = "" Then strMin = rsCopy!记录内容 & ""
                                            Exit For
                                        End If
                                        If (strMin <> "" Or strMax <> "") Then
                                            If strMin = strMax Then
                                                varItem(k - 1) = strMin
                                            ElseIf strMin <> strMax Then
                                                varItem(k - 1) = strMin & "-" & strMax
                                            End If
                                            strMin = "": strMax = ""
                                        End If
                                    End If
                                    mrs汇总时段.MoveNext
                                Next
                            ElseIf bytItem = 2 Then
                                '汇总项目
                                For k = 1 To mrs汇总时段.RecordCount
                                    If mrs汇总时段!开始 & "" <= mrs汇总时段!结束 & "" Then
                                        strFreB = Format(Format(datCurr, "YYYY-MM-DD ") & mrs汇总时段!开始, "YYYY-MM-DD HH:MM:SS")
                                        strFreE = Format(Format(datCurr, "YYYY-MM-DD ") & mrs汇总时段!结束, "YYYY-MM-DD HH:MM:SS")
                                    Else
                                        strFreB = Format(Format(datCurr, "YYYY-MM-DD ") & mrs汇总时段!开始, "YYYY-MM-DD HH:MM:SS")
                                        strFreE = Format(Format(DateAdd("d", 1, datCurr), "YYYY-MM-DD ") & mrs汇总时段!结束, "YYYY-MM-DD HH:MM:SS")
                                    End If
                                    If varItem(k - 1) = "" Then
                                          If Between(Format(rsCopy!发生时间 & "", "YYYY-MM-DD HH:MM:SS"), strFreB, strFreE) Then
                                            strMax = Val(strMax) + Val(rsCopy!记录内容 & "")
                                            Exit For
                                        End If
                                        If strMax <> "" Then
                                            varItem(k - 1) = strMax
                                            strMax = ""
                                        End If
                                    End If
                                    mrs汇总时段.MoveNext
                                Next
                             End If
                        End If
                        rsCopy.MoveNext
                    Loop
                    '将当天护理数据加入到数据集中
                    If rsCopy.RecordCount > 0 Then
                        If bytItem = 1 And (strMin <> "" Or strMax <> "") Then
                            '(strMin <> "" Or strMax <> "") 此条件用于数据时间不在汇总时段,导致K值等于3导致下标越界
                            If varItem(k - 1) = "" Then
                                If strMin = strMax Then
                                    varItem(k - 1) = strMin
                                ElseIf strMin <> strMax Then
                                    varItem(k - 1) = strMin & "-" & strMax
                                End If
                                strMin = "": strMax = ""
                            End If
                            If Val(mstr显示当天) = 0 Then
                                strPrveTime = DateAdd("d", 1, datCurr)
                            End If
                        ElseIf bytItem = 2 And strMax <> "" Then
                            If varItem(k - 1) = "" Then
                                varItem(k - 1) = strMax
                                strMax = ""
                            End If
                            If Val(mstr显示当天) = 0 Then
                                strPrveTime = DateAdd("d", 1, datCurr)
                            End If
                        End If
                        
                        strItem = "": If arrTmp(1) = "5" Then strTemp = colblood(Format(datCurr, "YYYY-MM-DDTHH:MM:SS"))
                        For k = LBound(varItem) To UBound(varItem)
                            If arrTmp(1) = "5" Then
                                strItem = strItem & "," & IIf(Split(strTemp, ",")(k) <> "" Or varItem(k) <> "", Split(strTemp, ",")(k) & "/" & varItem(k), Split(strTemp, ",")(k) & varItem(k))
                            Else
                                strItem = strItem & "," & varItem(k)
                            End If
                        Next
                        strItem = Mid(strItem, 2)
                        If arrTmp(1) = "4" Then
                            colblood.Add strItem, Format(datCurr, "YYYY-MM-DDTHH:MM:SS")
                        Else
                            With udtDataInfo
                                .Value = strItem
                                .Time = Format(datCurr, "YYYY-MM-DDTHH:MM:SS")
                                .Tag = strItem
                                colTmp.Add udtDataInfo, "_" & (colTmp.Count + 1)
                            End With
                        End If
                    End If
                    datCurr = DateAdd("d", 1, datCurr)
                Loop
            End If
        End If
        
        If arrTmp(1) = "4" Then arrTmp(1) = "5": GoTo LineBlood '血压:舒张压/收缩压
        
        If i < lngRow Then strTemp = mudtTimeLine.colFooterData("K_护理项目").ListData(i + 1).ItemTag
        Set udtDataItem.ListData = colTmp
        mudtTimeLine.colFooterData("K_护理项目").ListData.Remove (arrTmp(0))
        If i < lngRow Then
            mudtTimeLine.colFooterData("K_护理项目").ListData.Add udtDataItem, arrTmp(0), Split(strTemp, ",")(0)
        Else
            mudtTimeLine.colFooterData("K_护理项目").ListData.Add udtDataItem, arrTmp(0)
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

