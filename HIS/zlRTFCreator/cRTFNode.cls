VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFNode"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'全局公共类型定义应该在一个全局公共类中定义：


'图片类型
Public Enum PicTypeEnum
    pic_JPG = 0
    pic_GIF = 1
    pic_BMP = 2
    pic_PNG = 3
End Enum

'固定颜色表：
Public Enum ColorEnum
    clr00_默认 = 0
    clr01_黑色 = 1
    clr02_蓝色 = 2
    clr03_亮青 = 3
    clr04_绿色 = 4
    clr05_粉红 = 5
    clr06_大红 = 6
    clr07_黄色 = 7
    clr08_白色 = 8
    clr09_深蓝 = 9
    clr10_深青 = 10
    clr11_深绿 = 11
    clr12_浅棕 = 12
    clr13_深棕 = 13
    clr14_土黄 = 14
    clr15_深灰 = 15
    clr16_浅灰 = 16
End Enum

'字体尺寸：
Public Enum FontSizeEnum
    fs00_默认 = 0
    fs01_初号 = 42
    fs02_小初 = 36
    fs03_一号 = 26
    fs04_小一 = 24
    fs05_二号 = 22
    fs06_小二 = 18
    fs07_三号 = 16
    fs08_小三 = 15
    fs09_四号 = 14
    fs10_小四 = 12
    fs11_五号 = 10.5
    fs12_小五 = 9
    fs13_六号 = 7.5
    fs14_小六 = 6.5
    fs15_七号 = 5.5
    fs16_八号 = 5
End Enum

'字体样式：
Public Enum FontStyleEnum
    ft0_正文 = 0
    ft1_标题1 = 1
    ft2_标题2 = 2
    ft3_标题3 = 3
End Enum

'字符样式：
Public Enum CharStyleEnum
    cs默认 = 0
End Enum

'动态文本：
Public Enum AnimTypeEnum
    ani0_无 = 0
    ani1_七彩霓虹 = 1
    ani2_闪烁背景 = 2
    ani3_礼花绽放 = 3
    ani4_乌龙纹柱 = 4
    ani5_赤水情深 = 5
    ani6_亦真亦幻 = 6
End Enum

'浮雕雕刻
Public Enum EmboEnum
    emb0_无 = 0
    emb1_浮雕效果 = 1
    emb2_雕刻效果 = 2
End Enum

'下划线类型
Public Enum UnderLineTypeEnum
    udl00_无 = 0
    udl01_连续线 = 1
    udl02_点线 = 2
    udl03_短划线 = 3
    udl04_点划线 = 4
    udl05_双点划线 = 5
    udl06_双线 = 6
    udl07_加重波浪线 = 7
    udl08_长划线 = 8
    udl09_粗线 = 9
    udl10_粗点线 = 10
    udl11_粗短划线 = 11
    udl12_粗点划线 = 12
    udl13_粗双点划线 = 13
    udl14_粗长划线 = 14
    udl15_双波浪线 = 15
    udl16_字下加线 = 16
    udl17_波浪线 = 17
End Enum

'背景图案类型
Public Enum PatTypeEnum
    ptt00_无 = 0
    ptt01_横线 = 1
    ptt02_竖线 = 2
    ptt03_下斜线 = 3
    ptt04_上斜线 = 4
    ptt05_网格线 = 5
    ptt06_斜网格线 = 6
    ptt07_粗横线 = 7
    ptt08_粗竖线 = 8
    ptt09_粗下斜线 = 9
    ptt10_粗上斜线 = 10
    ptt11_粗网格线 = 11
    ptt12_粗斜网格线 = 12
End Enum

'对齐方式
Public Enum AlignTypeEnum
    alm0_默认对齐 = 0
    alm1_左端对齐 = 1
    alm2_居中对齐 = 2
    alm3_居右对齐 = 3
    alm4_两端对齐 = 4
    alm5_分散对齐 = 5
End Enum

'项目符号类型
Public Enum ListTypeEnum
    ltp00_无 = 0
    ltp01_文本1 = 1
    ltp02_文本2 = 2
    ltp03_文本3 = 3
    ltp04_文本4 = 4
    ltp05_文本5 = 5
    ltp06_文本6 = 6
    ltp07_文本7 = 7
    ltp08_文本8 = 8
    ltp09_符号1 = 9
    ltp10_符号2 = 10
    ltp11_符号3 = 11
    ltp12_符号4 = 12
    ltp13_符号5 = 13
    ltp14_符号6 = 14
    ltp15_符号7 = 15
End Enum


'垂直对齐方式
Public Enum AlignVTypeEnum
    alv0_默认对齐 = 0
    alv1_顶端对齐 = 1
    alv2_居中对齐 = 2
    alv3_底端对齐 = 3
End Enum

'单元格边框样式
Public Enum BorderStyleEnum
    cbs00_默认边框样式 = 0
    cbs01_无边框 = 1
    cbs02_单倍厚度边框 = 2
    cbs03_双倍厚度边框 = 3
    cbs04_阴影边框 = 4
    cbs05_双边框 = 5
    cbs06_点线边框 = 6
    cbs07_虚线边框 = 7
    cbs08_细线边框 = 8
    cbs09_插入式边框 = 9
    cbs10_虚线细边框 = 10
    cbs11_点划线边框 = 11
    cbs12_点点划线边框 = 12
    cbs13_初始边框 = 13
    cbs14_三重边框 = 14
    cbs15_厚薄叠加细边框 = 15
    cbs16_薄厚叠加细边框 = 16
    cbs17_薄厚薄叠加细边框 = 17
    cbs18_厚薄叠加中边框 = 18
    cbs19_薄厚叠加中边框 = 19
    cbs20_薄厚薄叠加中边框 = 20
    cbs21_厚薄叠加粗边框 = 21
    cbs22_薄厚叠加粗边框 = 22
    cbs23_薄厚薄叠加粗边框 = 23
    cbs24_波浪线边框 = 24
    cbs25_双波浪线边框 = 25
    cbs26_条纹边框 = 26
    cbs27_浮雕边框 = 27
    cbs28_雕刻边框 = 28
End Enum

'单元格长度单位
Public Enum CellLengthUnitEnum
    clt0_默认 = 0
    clt1_自动 = 1
    clt2_百分比 = 2
    clt3_缇 = 3
End Enum

'单元格内容类型
Public Enum CellContentTypeEnum
    cct0_文本 = 0
    cct1_段落 = 1
    cct2_图片 = 2
    cct3_表格 = 3
End Enum


'行对齐方式
Public Enum RowAlignTypeEnum
    ral0_默认对齐 = 0
    ral1_左端对齐 = 1
    ral2_居中对齐 = 2
    ral3_居右对齐 = 3
End Enum

'文档绘图方向
Public Enum DrawingDirectionEnum
    dwd0_水平绘图 = 0
    dwd1_垂直绘图 = 1
End Enum

'页面尺寸
Public Enum PaperSizeEnum
    pps0_Default = 0
    pps1_Letter = 1
    pps2_Tabloid = 2
    pps3_Legal = 3
    pps4_A3 = 4
    pps5_A4 = 5
    pps6_A5 = 6
    pps7_B4 = 7
    pps8_B5 = 8
End Enum

'页面视图
Public Enum ViewKindEnum
    vkd0_无 = 0
    vkd1_页视图 = 1
    vkd2_大纲视图 = 2
    vkd3_主文档视图 = 3
    vkd4_正常视图 = 4
    vkd5_联机视图 = 5
End Enum

'保持属性值的局部变量
Private mvarSubID As Long
Private mvarNodeType As String

Public Property Let NodeType(ByVal vData As String)
Attribute NodeType.VB_Description = "字符串，表示节点类型。文本、图片、表格等。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.NodeType = 5
    mvarNodeType = vData
End Property


Public Property Get NodeType() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.NodeType
    NodeType = mvarNodeType
End Property



Public Property Let SubID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.SubID = 5
    mvarSubID = vData
End Property


Public Property Get SubID() As Long
Attribute SubID.VB_Description = "子ID。文本、图片或者表格的ID。"
Attribute SubID.VB_UserMemId = 0
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SubID
    SubID = mvarSubID
End Property

