VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmProShow 
   BorderStyle     =   0  'None
   ClientHeight    =   8115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer timRest 
      Interval        =   8000
      Left            =   7560
      Top             =   6480
   End
   Begin VB.Timer timData 
      Interval        =   1000
      Left            =   8520
      Top             =   6480
   End
   Begin VB.Timer timTime 
      Interval        =   60000
      Left            =   9600
      Top             =   6480
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf待取药 
      Height          =   4455
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   5385
      _cx             =   9499
      _cy             =   7858
      Appearance      =   1
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   -2147483643
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   600
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmProShow.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.PictureBox picDraw 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   3240
      ScaleHeight     =   1035
      ScaleWidth      =   1800
      TabIndex        =   1
      Top             =   7320
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf已过号 
      Height          =   4455
      Left            =   7080
      TabIndex        =   2
      Top             =   1560
      Width           =   4185
      _cx             =   7382
      _cy             =   7858
      Appearance      =   1
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   -2147483643
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   600
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmProShow.frx":0152
      ScrollTrack     =   0   'False
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lbl温馨提示 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "温馨提示：请耐心等候，依次排队。祝您早日康复！"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   5
      Top             =   6240
      Width           =   5520
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2019-01-01 00:00"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7560
      TabIndex        =   10
      Top             =   7080
      Width           =   1965
   End
   Begin VB.Label lbl呼叫_窗口 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "一号窗口"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7800
      TabIndex        =   9
      Top             =   600
      Width           =   960
   End
   Begin VB.Label lbl呼叫_到 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "到"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7320
      TabIndex        =   8
      Top             =   600
      Width           =   240
   End
   Begin VB.Label lbl呼叫_姓名 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "陈无明"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6240
      TabIndex        =   7
      Top             =   600
      Width           =   720
   End
   Begin VB.Label lbl呼叫_请 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      TabIndex        =   6
      Top             =   600
      Width           =   240
   End
   Begin VB.Label lbl呼叫_取药 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "取药"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9240
      TabIndex        =   4
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lbl药房名称 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "门诊西药房"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   1200
   End
   Begin VB.Image img背景 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1065
   End
End
Attribute VB_Name = "frmProShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng药房ID As Long          '当前药房ID
Private mstrWins As String          '当前发药窗口
Private mbln配药 As Boolean         '是否配药
Private mbln配药确认 As Boolean     '是否配药确认
Private mbln正在显示 As Boolean     '呼叫内容是否正在显示

Private Const CST_STR_REG As String = "公共模块\药房排队叫号\液晶电视Pro"

Private Type Type_para
    bln单窗口模式 As Boolean          '屏幕窗口显示模式。True-单窗口显示；False-多窗口显示
    str多窗口 As String               '多窗口模式下所设置的各窗口
    
    str背景图片地址 As String         '本地背景图片访问地址
    
    '主窗体位置
    lngLeft As Long                  '主窗体距显示屏左边位置
    lngTop  As Long                 '主窗体距显示屏上边位置
    lngWidth As Long                '主窗体宽度
    lngHeight As Long               '主窗体高度
    
    '数据刷新
    int数据轮询时间 As Integer
    int呼叫显示时间 As Integer
    
    '药房相关设置
    lng药房_Left As Long
    lng药房_Top As Long
    
    str药房_字体 As String
    str药房_字号 As String
    bln药房_粗体 As Boolean
    bln药房_斜体 As Boolean
    lng药房_颜色 As Long
    
    '呼叫内容相关设置
    lng呼叫_Left As Long
    lng呼叫_Top As Long
    
    str呼叫通用_字体 As String
    str呼叫通用_字号 As String
    bln呼叫通用_粗体 As Boolean
    bln呼叫通用_斜体 As Boolean
    lng呼叫通用_颜色 As Long
    
    bln呼叫姓名单独设置 As Boolean
    str呼叫姓名_字体 As String
    str呼叫姓名_字号 As String
    bln呼叫姓名_粗体 As Boolean
    bln呼叫姓名_斜体 As Boolean
    lng呼叫姓名_颜色 As Long
    
    bln呼叫窗口单独设置 As Boolean
    str呼叫窗口_字体 As String
    str呼叫窗口_字号 As String
    bln呼叫窗口_粗体 As Boolean
    bln呼叫窗口_斜体 As Boolean
    lng呼叫窗口_颜色 As Long
    
    '待发药区域相关设置
    bln显示待发药 As Boolean
    
    lng待发药_Left As Long
    lng待发药_Top As Long
    
    str待发药_字体 As String
    str待发药_字号 As String
    bln待发药_粗体 As Boolean
    bln待发药_斜体 As Boolean
    lng待发药_颜色 As Long
    
    lng待发药_列宽 As Long
    lng待发药_行高 As Long
    lng待发药_行数 As Long
    
    '已过号区域相关设置
    bln显示已过号 As Boolean
    
    lng已过号_Left As Long
    lng已过号_Top As Long
    
    str已过号_字体 As String
    str已过号_字号 As String
    bln已过号_粗体 As Boolean
    bln已过号_斜体 As Boolean
    lng已过号_颜色 As Long
    
    lng已过号_列宽 As Long
    lng已过号_行高 As Long
    lng已过号_行数 As Long
    
    '温馨提示内容
    bln显示提示 As Boolean
    lng温馨提示_Left As Long
    lng温馨提示_Top As Long
    
    str温馨提示_内容 As String             '显示屏下方的温馨提示内容（默认指下方区域）
    
    str温馨提示_字体 As String
    str温馨提示_字号 As String
    bln温馨提示_粗体 As Boolean
    bln温馨提示_斜体 As Boolean
    lng温馨提示_颜色 As Long
    
    '时间
    bln显示时间 As Boolean
    
    lng时间_Left As Long
    lng时间_Top As Long
    
    str时间_字体 As String
    str时间_字号 As String
    bln时间_粗体 As Boolean
    bln时间_斜体 As Boolean
    lng时间_颜色 As Long
End Type

Private mPara As Type_para

Public Sub ShowMe(ByVal lng药房ID As Long, ByVal strWins As String, ByVal bln配药 As Boolean, ByVal bln配药确认 As Boolean)
    '功能：对模块进行初始化
    
    mlng药房ID = lng药房ID
    mstrWins = strWins
    mbln配药 = bln配药
    mbln配药确认 = bln配药确认
    
    '窗体显示
    Me.Show
End Sub

Private Sub Change呼叫内容()
    '功能：更新显示屏的呼叫内容
    Dim strSQL As String
    Dim strWins As String
    Dim date开始日期 As Date
    Dim date结束日期 As Date

    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mbln正在显示 Then Exit Sub
    
    strWins = IIf(mPara.bln单窗口模式, mstrWins, mPara.str多窗口)
    
    date开始日期 = gobjDatabase.Currentdate
    date开始日期 = CDate(Format(date开始日期, "yyyy-mm-dd") & " 00:00:00")

    date结束日期 = gobjDatabase.Currentdate
    date结束日期 = CDate(Format(date结束日期, "yyyy-mm-dd") & " 23:59:59")
    
    strSQL = "Select 姓名, 发药窗口, NO, 单据, 库房id" & vbNewLine & _
            "From (Select 姓名, 发药窗口, NO, 单据, 库房id" & vbNewLine & _
            "       From 未发药品记录" & vbNewLine & _
            "       Where (排队状态 = 3 Or 排队状态 = 4) And Nvl(显示状态, 0) = 0 And 库房id = [1]" & vbNewLine & _
            "       And 发药窗口 In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) And 填制日期 Between [3] And [4]" & vbNewLine & _
            "       Order By 呼叫时间) A" & vbNewLine & _
            "Where Rownum < 2"
        
    Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "加载正呼叫数据", mlng药房ID, strWins, date开始日期, date结束日期)
    
    If rsData.EOF Then Exit Sub
    
    lbl呼叫_姓名.Caption = rsData!姓名
    lbl呼叫_窗口.Caption = rsData!发药窗口
    
    mbln正在显示 = True
    
    timRest.Enabled = False
    timRest.Enabled = True
    
    Call Init呼叫内容
    Call Show呼叫内容(True)
    
    '呼叫完后清除呼叫内容，放在刷新显示的处理后面
    strSQL = "Zl_未发药品记录_显示("
    'NO
    strSQL = strSQL & "'" & rsData!NO & "'"
    '单据
    strSQL = strSQL & "," & rsData!单据
    '药房id
    strSQL = strSQL & "," & rsData!库房id
    strSQL = strSQL & ")"
    
    Call gobjDatabase.ExecuteProcedure(strSQL, "tmrCall_Timer")
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Public Sub Reset()
    '功能：重置界面布局
    
    Call LoadPar
    Call ResetLayout
    Call DrawWallPaper
End Sub

Private Sub LoadPar()
    '功能：获取参数设置
    
    With mPara
        .bln单窗口模式 = (Val(GetSetting("ZLSOFT", CST_STR_REG, "窗口模式", "0")) = 0)
        .str多窗口 = GetSetting("ZLSOFT", CST_STR_REG, "多窗口", "")
        .str背景图片地址 = GetSetting("ZLSOFT", CST_STR_REG, "图片位置", "")
        
        '主窗体位置
        .lngLeft = GetSetting("ZLSOFT", CST_STR_REG, "液晶屏_左", "1024")
        .lngTop = GetSetting("ZLSOFT", CST_STR_REG, "液晶屏_顶", "0")
        .lngWidth = GetSetting("ZLSOFT", CST_STR_REG, "液晶屏_宽度", "1024")
        .lngHeight = GetSetting("ZLSOFT", CST_STR_REG, "液晶屏_高度", "768")
        
        '数据刷新
        .int数据轮询时间 = Val(GetSetting("ZLSOFT", CST_STR_REG, "数据轮询时间", "1"))
        If .int数据轮询时间 < 1 Then
            .int数据轮询时间 = 1
        ElseIf .int数据轮询时间 > 60 Then
            .int数据轮询时间 = 60
        End If
        
        '显示刷新
        .int呼叫显示时间 = Val(GetSetting("ZLSOFT", CST_STR_REG, "呼叫显示时间", "1"))
        If .int呼叫显示时间 < 1 Then
            .int呼叫显示时间 = 1
        ElseIf .int呼叫显示时间 > 60 Then
            .int呼叫显示时间 = 60
        End If
        
        '药房相关设置
        .lng药房_Left = Val(GetSetting("ZLSOFT", CST_STR_REG, "药房_左", "0"))
        .lng药房_Top = Val(GetSetting("ZLSOFT", CST_STR_REG, "药房_顶", "0"))
        
        .str药房_字体 = GetSetting("ZLSOFT", CST_STR_REG, "药房字体", "微软雅黑")
        .str药房_字号 = GetSetting("ZLSOFT", CST_STR_REG, "药房字号", "12")
        .bln药房_粗体 = GetSetting("ZLSOFT", CST_STR_REG, "药房粗体", "False")
        .bln药房_斜体 = GetSetting("ZLSOFT", CST_STR_REG, "药房斜体", "False")
        .lng药房_颜色 = GetSetting("ZLSOFT", CST_STR_REG, "药房颜色", vbBlack)
        
        '呼叫内容相关设置
        .lng呼叫_Left = Val(GetSetting("ZLSOFT", CST_STR_REG, "呼叫_左", "0"))
        .lng呼叫_Top = Val(GetSetting("ZLSOFT", CST_STR_REG, "呼叫_顶", "0"))
        
        .str呼叫通用_字体 = GetSetting("ZLSOFT", CST_STR_REG, "呼叫字体_通用", "微软雅黑")
        .str呼叫通用_字号 = GetSetting("ZLSOFT", CST_STR_REG, "呼叫通用字体_字号", "12")
        .bln呼叫通用_粗体 = GetSetting("ZLSOFT", CST_STR_REG, "呼叫通用字体_粗体", "False")
        .bln呼叫通用_斜体 = GetSetting("ZLSOFT", CST_STR_REG, "呼叫通用字体_斜体", "False")
        .lng呼叫通用_颜色 = GetSetting("ZLSOFT", CST_STR_REG, "呼叫颜色_通用", vbBlack)
        
        .bln呼叫姓名单独设置 = (Val(GetSetting("ZLSOFT", CST_STR_REG, "呼叫姓名单独设置", "0")) = 1)
        .str呼叫姓名_字体 = GetSetting("ZLSOFT", CST_STR_REG, "呼叫字体_姓名", "微软雅黑")
        .str呼叫姓名_字号 = GetSetting("ZLSOFT", CST_STR_REG, "呼叫姓名字体_字号", "12")
        .bln呼叫姓名_粗体 = GetSetting("ZLSOFT", CST_STR_REG, "呼叫姓名字体_粗体", "False")
        .bln呼叫姓名_斜体 = GetSetting("ZLSOFT", CST_STR_REG, "呼叫姓名字体_斜体", "False")
        .lng呼叫姓名_颜色 = GetSetting("ZLSOFT", CST_STR_REG, "呼叫颜色_姓名", vbBlack)
        
        .bln呼叫窗口单独设置 = (Val(GetSetting("ZLSOFT", CST_STR_REG, "呼叫窗口单独设置", "0")) = 1)
        .str呼叫窗口_字体 = GetSetting("ZLSOFT", CST_STR_REG, "呼叫字体_窗口", "微软雅黑")
        .str呼叫窗口_字号 = GetSetting("ZLSOFT", CST_STR_REG, "呼叫窗口字体_字号", "12")
        .bln呼叫窗口_粗体 = GetSetting("ZLSOFT", CST_STR_REG, "呼叫窗口字体_粗体", "False")
        .bln呼叫窗口_斜体 = GetSetting("ZLSOFT", CST_STR_REG, "呼叫窗口字体_斜体", "False")
        .lng呼叫窗口_颜色 = GetSetting("ZLSOFT", CST_STR_REG, "呼叫颜色_窗口", vbBlack)
        
        '待发药区域相关设置
        .bln显示待发药 = (Val(GetSetting("ZLSOFT", CST_STR_REG, "显示待发药", "1")) = 1)
        
        .lng待发药_Left = Val(GetSetting("ZLSOFT", CST_STR_REG, "待发药_左", "0"))
        .lng待发药_Top = Val(GetSetting("ZLSOFT", CST_STR_REG, "待发药_顶", "0"))
        
        .str待发药_字体 = GetSetting("ZLSOFT", CST_STR_REG, "待发药字体", "微软雅黑")
        .str待发药_字号 = GetSetting("ZLSOFT", CST_STR_REG, "待发药字号", "12")
        .bln待发药_粗体 = GetSetting("ZLSOFT", CST_STR_REG, "待发药粗体", "False")
        .bln待发药_斜体 = GetSetting("ZLSOFT", CST_STR_REG, "待发药斜体", "False")
        .lng待发药_颜色 = GetSetting("ZLSOFT", CST_STR_REG, "待发药颜色", vbBlack)
        
        .lng待发药_列宽 = Val(GetSetting("ZLSOFT", CST_STR_REG, "待发药_列宽", "800"))
        .lng待发药_行高 = Val(GetSetting("ZLSOFT", CST_STR_REG, "待发药_行高", "350"))
        .lng待发药_行数 = Val(GetSetting("ZLSOFT", CST_STR_REG, "待发药_行数", "5"))
        
        '已过号区域相关设置
        .bln显示已过号 = (Val(GetSetting("ZLSOFT", CST_STR_REG, "显示已过号", "1")) = 1)
        
        .lng已过号_Left = Val(GetSetting("ZLSOFT", CST_STR_REG, "已过号_左", "0"))
        .lng已过号_Top = Val(GetSetting("ZLSOFT", CST_STR_REG, "已过号_顶", "0"))
        
        .str已过号_字体 = GetSetting("ZLSOFT", CST_STR_REG, "已过号字体", "微软雅黑")
        .str已过号_字号 = GetSetting("ZLSOFT", CST_STR_REG, "已过号字号", "12")
        .bln已过号_粗体 = GetSetting("ZLSOFT", CST_STR_REG, "已过号粗体", "False")
        .bln已过号_斜体 = GetSetting("ZLSOFT", CST_STR_REG, "已过号斜体", "False")
        .lng已过号_颜色 = GetSetting("ZLSOFT", CST_STR_REG, "已过号颜色", vbBlack)
        
        .lng已过号_列宽 = Val(GetSetting("ZLSOFT", CST_STR_REG, "已过号_列宽", "800"))
        .lng已过号_行高 = Val(GetSetting("ZLSOFT", CST_STR_REG, "已过号_行高", "350"))
        .lng已过号_行数 = Val(GetSetting("ZLSOFT", CST_STR_REG, "已过号_行数", "5"))
        
        '温馨提示内容
        .bln显示提示 = (Val(GetSetting("ZLSOFT", CST_STR_REG, "显示提示", "1")) = 1)
        .lng温馨提示_Left = Val(GetSetting("ZLSOFT", CST_STR_REG, "提示_左", "0"))
        .lng温馨提示_Top = Val(GetSetting("ZLSOFT", CST_STR_REG, "提示_顶", "0"))
        
        .str温馨提示_内容 = GetSetting("ZLSOFT", CST_STR_REG, "提示_内容", "")
        
        .str温馨提示_字体 = GetSetting("ZLSOFT", CST_STR_REG, "提示字体", "微软雅黑")
        .str温馨提示_字号 = GetSetting("ZLSOFT", CST_STR_REG, "提示字号", "12")
        .bln温馨提示_粗体 = GetSetting("ZLSOFT", CST_STR_REG, "提示粗体", "False")
        .bln温馨提示_斜体 = GetSetting("ZLSOFT", CST_STR_REG, "提示斜体", "False")
        .lng温馨提示_颜色 = GetSetting("ZLSOFT", CST_STR_REG, "提示颜色", vbBlack)
        
        '时间
        .bln显示时间 = (Val(GetSetting("ZLSOFT", CST_STR_REG, "显示时间", "1")) = 1)
        
        .lng时间_Left = Val(GetSetting("ZLSOFT", CST_STR_REG, "时间_左", "0"))
        .lng时间_Top = Val(GetSetting("ZLSOFT", CST_STR_REG, "时间_顶", "0"))
        
        .str时间_字体 = GetSetting("ZLSOFT", CST_STR_REG, "时间字体", "微软雅黑")
        .str时间_字号 = GetSetting("ZLSOFT", CST_STR_REG, "时间字号", "12")
        .bln时间_粗体 = GetSetting("ZLSOFT", CST_STR_REG, "时间粗体", "False")
        .bln时间_斜体 = GetSetting("ZLSOFT", CST_STR_REG, "时间斜体", "False")
        .lng时间_颜色 = GetSetting("ZLSOFT", CST_STR_REG, "时间颜色", vbBlack)
            
    End With
    
End Sub

Private Sub Init呼叫内容()
    '功能：重置呼叫内容的位置及样式
    
    With lbl呼叫_请
        .FontName = mPara.str呼叫通用_字体
        .FontSize = mPara.str呼叫通用_字号
        .FontBold = mPara.bln呼叫通用_粗体
        .FontItalic = mPara.bln呼叫通用_斜体
        .ForeColor = mPara.lng呼叫通用_颜色
        
        .Left = mPara.lng呼叫_Left
        .Top = mPara.lng呼叫_Top
    End With
    
    With lbl呼叫_姓名
        .FontName = IIf(mPara.bln呼叫姓名单独设置, mPara.str呼叫姓名_字体, mPara.str呼叫通用_字体)
        .FontSize = IIf(mPara.bln呼叫姓名单独设置, mPara.str呼叫姓名_字号, mPara.str呼叫通用_字号)
        .FontBold = IIf(mPara.bln呼叫姓名单独设置, mPara.bln呼叫姓名_粗体, mPara.bln呼叫通用_粗体)
        .FontItalic = IIf(mPara.bln呼叫姓名单独设置, mPara.bln呼叫姓名_斜体, mPara.bln呼叫通用_斜体)
        .ForeColor = IIf(mPara.bln呼叫姓名单独设置, mPara.lng呼叫姓名_颜色, mPara.lng呼叫通用_颜色)
        
        .Left = lbl呼叫_请.Left + lbl呼叫_请.Width + 50
        .Top = lbl呼叫_请.Top + (lbl呼叫_请.Height - .Height) / 2
    End With
    
    With lbl呼叫_到
        .FontName = mPara.str呼叫通用_字体
        .FontSize = mPara.str呼叫通用_字号
        .FontBold = mPara.bln呼叫通用_粗体
        .FontItalic = mPara.bln呼叫通用_斜体
        .ForeColor = mPara.lng呼叫通用_颜色
        
        .Left = lbl呼叫_姓名.Left + lbl呼叫_姓名.Width + 50
        .Top = lbl呼叫_请.Top
    End With
    
    With lbl呼叫_窗口
        .FontName = IIf(mPara.bln呼叫窗口单独设置, mPara.str呼叫窗口_字体, mPara.str呼叫通用_字体)
        .FontSize = IIf(mPara.bln呼叫窗口单独设置, mPara.str呼叫窗口_字号, mPara.str呼叫通用_字号)
        .FontBold = IIf(mPara.bln呼叫窗口单独设置, mPara.bln呼叫窗口_粗体, mPara.bln呼叫通用_粗体)
        .FontItalic = IIf(mPara.bln呼叫窗口单独设置, mPara.bln呼叫窗口_斜体, mPara.bln呼叫通用_斜体)
        .ForeColor = IIf(mPara.bln呼叫窗口单独设置, mPara.lng呼叫窗口_颜色, mPara.lng呼叫通用_颜色)
        
        .Left = lbl呼叫_到.Left + lbl呼叫_到.Width + 50
        .Top = lbl呼叫_到.Top + (lbl呼叫_到.Height - .Height) / 2
    End With
    
    With lbl呼叫_取药
        .FontName = mPara.str呼叫通用_字体
        .FontSize = mPara.str呼叫通用_字号
        .FontBold = mPara.bln呼叫通用_粗体
        .FontItalic = mPara.bln呼叫通用_斜体
        .ForeColor = mPara.lng呼叫通用_颜色
        
        .Left = lbl呼叫_窗口.Left + lbl呼叫_窗口.Width + 50
        .Top = lbl呼叫_请.Top
    End With
End Sub

Private Sub ResetLayout()
    '窗体位置
    With Me
        .Left = mPara.lngLeft * Screen.TwipsPerPixelY
        .Top = mPara.lngTop * Screen.TwipsPerPixelY
        .Width = mPara.lngWidth * Screen.TwipsPerPixelX
        .Height = mPara.lngHeight * Screen.TwipsPerPixelY
    End With
    
    '背景设置
    With img背景
        .Top = 0
        .Left = 0
        .Height = Me.ScaleHeight
        .Width = Me.ScaleWidth
        
        .Picture = LoadPicture(mPara.str背景图片地址)
    End With
    
    '药房相关设置
    With lbl药房名称
        .Left = mPara.lng药房_Left
        .Top = mPara.lng药房_Top
        
        .FontName = mPara.str药房_字体
        .FontSize = mPara.str药房_字号
        .FontBold = mPara.bln药房_粗体
        .FontItalic = mPara.bln药房_斜体
        .ForeColor = mPara.lng药房_颜色
    End With
    
    '数据轮询时间
    With timData
        .Interval = mPara.int数据轮询时间 * 1000#
    End With
    
    '呼叫显示时间
    With timRest
        .Interval = mPara.int呼叫显示时间 * 1000#
    End With
    
    '呼叫内容相关设置
    Call Init呼叫内容
    
    '待发药区域相关设置
    Call InitList_待发药
    
    '已过号区域相关设置
    Call InitList_已过号
    
    '温馨提示内容
    With lbl温馨提示
        .Visible = mPara.bln显示提示
        .Caption = mPara.str温馨提示_内容
        
        .FontName = mPara.str温馨提示_字体
        .FontSize = mPara.str温馨提示_字号
        .FontBold = mPara.bln温馨提示_粗体
        .FontItalic = mPara.bln温馨提示_斜体
        .ForeColor = mPara.lng温馨提示_颜色
        
        .Left = mPara.lng温馨提示_Left
        .Top = mPara.lng温馨提示_Top
    End With
    
    '时间
    With lblTime
        .Visible = mPara.bln显示时间
        
        .Caption = Format(gobjDatabase.Currentdate, "yyyy-mm-dd  hh:mm")
        
        .FontName = mPara.str时间_字体
        .FontSize = mPara.str时间_字号
        .FontBold = mPara.bln时间_粗体
        .FontItalic = mPara.bln时间_斜体
        .ForeColor = mPara.lng时间_颜色
        
        .Left = mPara.lng时间_Left
        .Top = mPara.lng时间_Top
    End With
    
    With timTime
        .Enabled = mPara.bln显示时间
    End With
    
End Sub

Private Sub Form_Load()
    mbln正在显示 = False
    
    Call Show呼叫内容(False)
        
    '获取参数设置
    Call LoadPar
    
    '重置布局
    Call ResetLayout
    
    '重绘表格背景
    Call DrawWallPaper
End Sub

Private Sub Show呼叫内容(ByVal bln显示 As Boolean)
    '功能：是否显示呼叫内容
    
    lbl呼叫_请.Visible = bln显示
    lbl呼叫_姓名.Visible = bln显示
    lbl呼叫_到.Visible = bln显示
    lbl呼叫_窗口.Visible = bln显示
    lbl呼叫_取药.Visible = bln显示
End Sub


Private Sub InitList_待发药()
    '功能：初始化待发药列表
    Dim str窗口串 As String
    Dim i As Integer
    
    str窗口串 = IIf(mPara.bln单窗口模式, mstrWins, mPara.str多窗口)
    
    With vsf待取药
        .Visible = mPara.bln显示待发药
        .Left = mPara.lng待发药_Left
        .Top = mPara.lng待发药_Top
        
        '数据清空
        .Clear
        
        '定义行数
        If mPara.bln单窗口模式 Then
            .Rows = 1 + mPara.lng待发药_行数
        Else
            .Rows = 4 + mPara.lng待发药_行数
        End If
        
        '定义表格的列数
        .Cols = UBound(Split(str窗口串, ",")) + 1
        
        '定义表格每列Key值
        For i = 0 To UBound(Split(str窗口串, ","))
            .ColKey(i) = Split(str窗口串, ",")(i)
            
            If mPara.bln单窗口模式 Then
                '显示“等待取药”
                .TextMatrix(0, i) = "等待取药"
            Else
                '显示“发药窗口”
                .TextMatrix(0, i) = Split(str窗口串, ",")(i)
                
                '显示“正在取药”
                .TextMatrix(1, i) = "当前取药"
                
                '显示“等待取药”
                .TextMatrix(3, i) = "等待取药"
            End If
        Next
        
        '居中显示
        .ColAlignment(-1) = flexAlignCenterCenter
        
        '字体设置
        .FontName = mPara.str待发药_字体
        .FontSize = mPara.str待发药_字号
        .FontBold = mPara.bln待发药_粗体
        .FontItalic = mPara.bln待发药_斜体
        .ForeColor = mPara.lng待发药_颜色
        
        '单元格设置
        .ColWidth(-1) = mPara.lng待发药_列宽
        .RowHeight(-1) = mPara.lng待发药_行高
        
        .Width = mPara.lng待发药_列宽 * .Cols
        .Height = mPara.lng待发药_行高 * .Rows
        
    End With
    
    '数据初始化
    Call LoadData_待发药
    
End Sub

Private Sub InitList_已过号()
    '功能：初始化待发药列表
    Dim str窗口串 As String
    Dim i As Integer
    
    str窗口串 = IIf(mPara.bln单窗口模式, mstrWins, mPara.str多窗口)
    
    With vsf已过号
        .Visible = mPara.bln显示已过号
        .Left = mPara.lng已过号_Left
        .Top = mPara.lng已过号_Top
        
        '数据清空
        .Clear
    
        '定义行数
        If mPara.bln单窗口模式 Then
            .Rows = 1 + mPara.lng待发药_行数
        Else
            .Rows = 2 + mPara.lng已过号_行数
        End If
        
        '定义表格的列数
        .Cols = UBound(Split(str窗口串, ",")) + 1
    
        '定义表格每列Key值
        For i = 0 To UBound(Split(str窗口串, ","))
            .ColKey(i) = Split(str窗口串, ",")(i)
            
            '显示“过号名单”
            .TextMatrix(0, i) = "过号名单"
            
            If Not mPara.bln单窗口模式 Then
                '显示“发药窗口”
                .TextMatrix(1, i) = Split(str窗口串, ",")(i)
            End If
            
        Next
        
        '居中显示
        .ColAlignment(-1) = flexAlignCenterCenter
        
        '合并
        .MergeCells = flexMergeRestrictRows
        .MergeRow(0) = True
        
        '字体设置
        .FontName = mPara.str已过号_字体
        .FontSize = mPara.str已过号_字号
        .FontBold = mPara.bln已过号_粗体
        .FontItalic = mPara.bln已过号_斜体
        .ForeColor = mPara.lng已过号_颜色
        
        '单元格设置
        .ColWidth(-1) = mPara.lng已过号_列宽
        .RowHeight(-1) = mPara.lng已过号_行高
        
        .Width = mPara.lng已过号_列宽 * .Cols
        .Height = mPara.lng已过号_行高 * .Rows
        
    End With
    
    '数据初始化
    Call LoadData_已过号
    
End Sub

Private Sub RefreshList_正呼叫(ByVal rsData As ADODB.Recordset)
    '功能：刷新待发药列表
    Dim i As Integer
    Dim n As Integer
    Dim int固定列 As Integer       '默认第二行为正呼叫信息
    
    int固定列 = 2
    
    With vsf待取药
        For i = 0 To .Cols - 1
            rsData.Filter = "发药窗口 = '" & .ColKey(i) & "'"
            
            '重写数据
            If Not rsData.EOF Then
                .TextMatrix(int固定列, i) = rsData!姓名
            Else
                .TextMatrix(int固定列, i) = ""
            End If
        Next
    End With
End Sub

Private Sub RefreshList_待发药(ByVal rsData As ADODB.Recordset)
    '功能：刷新待发药列表
    Dim i As Integer
    Dim n As Integer
    Dim int起步行 As Integer
    Dim intTemp As Integer
    
    int起步行 = IIf(mPara.bln单窗口模式, 0, 3)
    
    With vsf待取药
        For i = 0 To .Cols - 1
            rsData.Filter = "发药窗口 = '" & .ColKey(i) & "'"
            
            '暂时清空该列数据
            For n = int起步行 + 1 To .Rows - 1
                .TextMatrix(n, i) = ""
            Next
            
            intTemp = IIf(rsData.RecordCount > mPara.lng待发药_行数, mPara.lng待发药_行数, rsData.RecordCount)
            
            '重新加载该列数据
            For n = 1 To intTemp
                .TextMatrix(n + int起步行, i) = rsData!姓名
                
                rsData.MoveNext
            Next
        Next
    End With
End Sub

Private Sub RefreshList_已过号(ByVal rsData As ADODB.Recordset)
    '功能：刷新已过号列表
    Dim i As Integer
    Dim n As Integer
    Dim int起步行 As Integer
    Dim intTemp As Integer
    
    int起步行 = IIf(mPara.bln单窗口模式, 0, 1)
    
    With vsf已过号
        For i = 0 To .Cols - 1
            rsData.Filter = "发药窗口 = '" & .ColKey(i) & "'"
            
            '暂时清空该列数据
            For n = int起步行 + 1 To .Rows - 1
                .TextMatrix(n, i) = ""
            Next
            
            intTemp = IIf(rsData.RecordCount > mPara.lng已过号_行数, mPara.lng已过号_行数, rsData.RecordCount)
            
            '重新加载该列数据
            For n = 1 To intTemp
                .TextMatrix(n + int起步行, i) = rsData!姓名
                
                rsData.MoveNext
            Next
        Next
    End With
    
End Sub

Private Sub LoadData_正呼叫()
    '功能：加载正呼叫数据
    '入参：【strWins】传入的发药窗口
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    Dim date开始日期 As Date
    Dim date结束日期 As Date
    Dim strWins As String
    
    On Error GoTo errHandle
    
    If mPara.bln单窗口模式 Then Exit Sub
    
    strWins = mPara.str多窗口
    
    date开始日期 = gobjDatabase.Currentdate
    date开始日期 = CDate(Format(date开始日期, "yyyy-mm-dd") & " 00:00:00")

    date结束日期 = gobjDatabase.Currentdate
    date结束日期 = CDate(Format(date结束日期, "yyyy-mm-dd") & " 23:59:59")
    
    strSQL = "Select 姓名,发药窗口" & vbNewLine & _
            "From 未发药品记录" & vbNewLine & _
            "Where 排队状态 = 3 And 库房id = [1] And 发药窗口 In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) And 填制日期 Between [3] And [4]"

    Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "加载正呼叫数据", mlng药房ID, strWins, date开始日期, date结束日期)
    
    '刷新正呼叫界面数据
    Call RefreshList_正呼叫(rsData)
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub LoadData_待发药()
    '功能：加载待发药数据
    '入参：【strWins】传入的发药窗口
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    Dim date开始日期 As Date
    Dim date结束日期 As Date
    Dim strWins As String
    
    On Error GoTo errHandle
    
    strWins = IIf(mPara.bln单窗口模式, mstrWins, mPara.str多窗口)
    
    date开始日期 = gobjDatabase.Currentdate
    date开始日期 = CDate(Format(date开始日期, "yyyy-mm-dd") & " 00:00:00")

    date结束日期 = gobjDatabase.Currentdate
    date结束日期 = CDate(Format(date结束日期, "yyyy-mm-dd") & " 23:59:59")
    
    strSQL = "Select A.病人ID,A.姓名,B.配药日期,B.签到时间,B.填制日期,A.发药窗口 " & _
            "From 未发药品记录 A,药品收发记录 B " & _
            "Where A.单据=B.单据 And A.No=B.NO And A.库房id=B.库房id And (A.单据=8 or A.单据=9 or A.单据=10) "

    If mbln配药 Then
        strSQL = strSQL & " and A.排队状态=2 and A.库房id=[1] and A.发药窗口 In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0)"
    ElseIf mbln配药确认 And mbln配药 = False Then
        strSQL = strSQL & " and (A.排队状态=1 or A.排队状态=2) and A.库房id=[1] and A.发药窗口 In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0)"
    ElseIf mbln配药 = False And mbln配药确认 = False Then
        strSQL = strSQL & " and (A.排队状态<3 or A.排队状态 is null) and A.库房id=[1] and A.发药窗口 In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0)"
    End If
    
    strSQL = "Select Rownum 序号,姓名,日期,发药窗口 " & _
            "From ( " & _
            "Select min(" & IIf(mbln配药, "配药日期", "Nvl(签到时间,填制日期)") & ") 日期,病人id,姓名,发药窗口 " & _
            "From (" & strSQL & ") " & _
            "Where 病人ID Not In (Select distinct A.病人ID From 未发药品记录 A,药品收发记录 B,门诊费用记录 C " & _
            "Where A.单据=B.单据 And A.No=B.NO And A.库房id=B.库房id and B.费用id=C.id and (A.单据=8 or A.单据=9 or A.单据=10) " & _
            "  and (A.排队状态=4 or A.排队状态 = 3) and A.库房id=[1] and A.发药窗口 In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0)) " & _
            "Group By 姓名,病人id,发药窗口 " & _
            "Order by 日期 " & _
            ")"
                    
    Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "加载待发药数据", mlng药房ID, strWins, date开始日期, date结束日期)
    
    '刷新待发药界面数据
    Call RefreshList_待发药(rsData)
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub LoadData_已过号()
    '功能：加载已过号数据
    '入参：【strWins】传入的发药窗口
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim date开始日期 As Date
    Dim date结束日期 As Date
    Dim strWins As String
    
    On Error GoTo errHandle
    
    strWins = IIf(mPara.bln单窗口模式, mstrWins, mPara.str多窗口)
    
    date开始日期 = gobjDatabase.Currentdate
    date开始日期 = CDate(Format(date开始日期, "yyyy-mm-dd") & " 00:00:00")

    date结束日期 = gobjDatabase.Currentdate
    date结束日期 = CDate(Format(date结束日期, "yyyy-mm-dd") & " 23:59:59")
    
    strSQL = "Select distinct A.病人ID,A.姓名,A.呼叫时间, A.发药窗口 From 未发药品记录 A,药品收发记录 B,门诊费用记录 C" & _
            " Where A.单据=B.单据 And A.No=B.NO And A.库房id=B.库房id and B.费用id=C.id and (A.单据=8 or A.单据=9 or A.单据=10) " & _
            "   and A.排队状态=4 and A.库房id=[1] and A.发药窗口 In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0) " & _
            " And a.呼叫时间 < Sysdate - 1 / 24 / 60 / 60 * [5] "
            
    strTemp = Replace(strSQL, "门诊费用记录", "住院费用记录")
            
    strSQL = strSQL & " union all " & strTemp
    
    strSQL = "Select rownum 序号,姓名,发药窗口 " & _
            "From (Select 病人ID,姓名,min(呼叫时间) 呼叫时间,发药窗口 " & _
            "From (" & strSQL & ") " & _
            "Group by 病人ID,姓名,发药窗口 " & _
            "Order by 呼叫时间 asc " & _
            ")"
    
    Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "加载已过号数据", mlng药房ID, strWins, date开始日期, date结束日期, 1)
    
    '刷新已过号界面数据
    Call RefreshList_已过号(rsData)
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub DrawWallPaper()
    '功能：设置表格控件的背景图片为表格背后的区域图片
    Dim std As StdPicture
    Dim strTempFile As String
    
    On Error GoTo errHandle
    
    '待发药表格处理
    '--------------------------------
    picDraw.cls
    picDraw.Width = vsf待取药.Width
    picDraw.Height = vsf待取药.Height
    
    picDraw.PaintPicture img背景.Picture, -vsf待取药.Left, -vsf待取药.Top, Me.ScaleWidth, Me.ScaleHeight
    
    Set vsf待取药.WallPaper = picDraw.Image
    '--------------------------------
    
    '已过号表格处理
    '--------------------------------
    picDraw.cls
    picDraw.Width = vsf已过号.Width
    picDraw.Height = vsf已过号.Height
    
    picDraw.PaintPicture img背景.Picture, -vsf已过号.Left, -vsf已过号.Top, Me.ScaleWidth, Me.ScaleHeight
    
    Set vsf已过号.WallPaper = picDraw.Image
    '--------------------------------

    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub timData_Timer()
    '功能：定时刷新显示屏数据
    
    Call LoadData_正呼叫
    Call LoadData_待发药
    Call LoadData_已过号
    
    Call Change呼叫内容
End Sub

Private Sub timRest_Timer()
    mbln正在显示 = False
    
    Call Show呼叫内容(False)
End Sub

Private Sub timTime_Timer()
    '功能：刷新时间
    lblTime.Caption = Format(gobjDatabase.Currentdate, "yyyy-mm-dd  hh:mm")
End Sub
