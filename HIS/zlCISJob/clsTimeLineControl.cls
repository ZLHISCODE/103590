VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimeLineControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum E_TICK_VALUE
    TICK_1 = 1
    TICK_2 = 2
    TICK_3 = 3
    TICK_4 = 4
    TICK_6 = 6
    TICK_8 = 8
    TICK_12 = 12
    TICK_24 = 24
End Enum

Public Type DesignInfoTickRange
    DisplayValue As Integer                   '显示值
    TickStartTime As String                 '时刻开始时间第一个从0：0开始
End Type

'时间轴设计样式
Public Type TimeLineDesignInfo
    BackgroundColor As String                   '背景颜色 "200,255,255,255" A(透明度0-255),R(0-255),G(0-255),B(0-255)没有设置，默认颜色就是透明，字符串也可以用Transparent空字符串默认也是全透明
    DateTitle As String                         '日期行标题"日期"
    DateTitleColor  As String
    DateTitleFont  As String
    DateStart As String                          '开始日期，小于开始日期那天的数据不显示。 允许为空
    dateEnd   As String                          '结束日期，大于日期那天的数据不显示。 允许为空
    ShowTick As Boolean                           '是否显示时刻行     true
    ShowFullDate As Boolean                        '日期行显示完整日期
    TickTitle As String                         '时刻行标题 时刻
    TickTitleColor  As String
    TickTitleFont  As String
    TickWidth As Integer                          '时刻宽度   18
    TickRangeListCount As E_TICK_VALUE                '时刻范围间隔 取值范围{1,2,3,4,6,8,12,24}
    DateFont As String                  '日期字体<DateFont>宋体, 9pt</DateFont>
    TickFont As String                  '时刻字体<TickFont>宋体, 9pt</TickFont>
    MergePeriodWidth As Integer              '数据区域合并后显示宽度
    EmptyDataMergeDayCount As Integer        '多少天无数据合并
    EmptyDataMergePeriodWidth As Integer     '无数据合并后显示宽度
    PaddingLeft As Integer                    '内部左边距
    PaddingTop As Integer                    '内部上边距
    PaddingRight As Integer                  '内部右边距
    PaddingBottom As Integer                 '内部下边距
    RowHeight As Integer                     '默认行高

    MeasureTitleWidth As Integer             '标尺左边标题列宽度
    GridMinHeight As Integer                 '标尺数据网格最小高度
    TopFixedSmallRowCount As Integer         '标尺顶部不显示数据行数
    BottomFixedSmallRowCount As Integer      '标尺底部不显示数据行数
    GridYSplitCount As Integer               '标尺网格垂直分割数
    GridYSmallSplitCount As Integer          '标尺网格小格垂直分割数
    Height As Integer                        '标尺高度 (Null表示可伸缩)
End Type

Public Type DataInfo
    Value As String
    Time As String              '实例:2015-07-22T11:50:25
    RowNumber As Integer        '连续数据用
    TimeEnd As String           '实例:2015-07-22T11:50:25
    Tag  As String
                                '对于【TextColor】、【BackgroudColor】、【FontName】、【FontSize】节点，【DataInfo】默认继承【DataItem】的值，如果有设置就用自己的值。
    BackgroundColor As String
    TextColor As String         'Black
    Font As String          '字体相关属性<Font>宋体, 9.75pt</Font>
    RowIndex As String         '行号，作为数据唯一标识符
                                            '对于【LegendType】、【ShadowLegendType】、【BalloonLegendType】节点，【DataInfo】默认继承【DataItem】的值，如果有设置就用自己的值。
    LegendType As String                    '数据点显示符号{实心圆、空心圆、粗空心圆、点、叉、H符号}
    ShadowLegendType As String              '阴影点显示符号
    BalloonLegendType As String             '气球点显示符号

    NumberValue As String            '点的数据值
    ShadowValue As String           '阴影点的值<ShadowValue xsi:nil="true" />
    BalloonValue As String          '气球点的值
    Group  As String                '相同Group的数据前显示分组线。
    Tip As String                   '表头底部数据信息增加Tip节点，显示提示（计数数据除外），为空时提示Value节点的值。
    HotspotFont  As String                  '热点字体
    HotspotColor As String                  '热点颜色
    ShowHotspotCursor As Boolean            '热点指针
End Type

Public Type DataItem
    GraphType As String
    Title As String
    TitleColor As String        '左边标题支持颜色设置。
    TitleFont As String         '左边标题支持字体设置。
    BackgroundColor As String
    TextColor As String           '默认设置 Black
    Font As String            ' 字体相关属性<Font>宋体, 9.75pt</Font>
    ListData As Collection    'DataItem或DataInfo的集合
    
    SplitString As String          '用于分割数据
    SplitCount  As Integer
    Unit As String              '单位，显示在【Title】后面，超宽自动换行 用于标尺数据
    MinValue As Integer         '标尺刻度最小值
    MaxValue As Integer         '标尺刻度最大值
    SplitNum As Integer         '刻度分割数
    SplitScale As Integer       '自定义分割，有值时【SplitNum】无效，比如视力（0,1,2,3,4,4.3,4.6,4.9,5.2）。
    Color As String              '标尺点和线的颜色
    IsDataDynamicExpansion As Boolean     '是否动态扩展数据，动态计算位置，根据当前显示周期内数据确定刻度最大值和最小值
    ShadowTitle As String                  '阴影图形名称。
    BalloonTitle As String                  '气球图形名称
    LegendType As String                    '数据点显示符号{实心圆、空心圆、粗空心圆、点、叉、H符号}
    ShadowLegendType As String              '阴影点显示符号
    BalloonLegendType As String             '气球点显示符号
    ShadowColor As String                   '标尺DataItem节点  没有该值使用Color颜色
    BalloonColor As String                  '标尺DataItem节点  没有该值使用Color颜色
    StartDate As String                     '开始日期，大于等于所给值。
    EndDate As String                       '结束日期，小于所给值
    FixedTick As String                     '固定时刻
    EquantTick As String                    '等分时刻每个时刻的大小，【FixedTick】优先。
    EquantTickUnit As String                '等分时刻每个时刻的单位（小时、分钟）。
    TickWidth  As Integer                   '时刻的宽度
                                            '数据区域DataItem节点
    LineColor As String                      '网格线颜色
    IsCollapse As Boolean                   '是否折叠
    ShowHotspotEffect As Boolean              '热点效果
    ShowHotspotCursor As Boolean            '热点指针
    HotspotFont  As String                  '热点字体 <HotspotFont>宋体, 9.75pt, style=Bold, Underline</HotspotFont>
    HotspotColor As String                  '热点颜色
    BorderColor As String                   '排版文本边框颜色
    GroupPosition  As String                '分组位置:居左,居上  默认居左
    Effect  As String                       '取值(无,网格）
    
    ItemTag As String                       '用于缓存数据标识,非时间轴控件节点值
End Type

Public Type TimeLineData
    colHeaderData As Collection                 '表头
    colFooterData As Collection                 '页脚数据
    colMeasureData As Collection                '标尺数据
    colMeasureVerticalText As Collection        '垂直文本
    colCustomTick As Collection                 '时刻自定义
    colDataArea  As Collection                  '数据区域自定义
End Type

