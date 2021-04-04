Attribute VB_Name = "mdlSystemCortrol"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''[图像保留的“系统标注”说明]''''''''''''''''''''''''''''''''''
''''图像内部保留了50个系统标注。
''''同时还有几种特殊类型的标注，LabelType = 7（体位标注doLabelSpecial）和LabelType = 11（标尺doLabelRuler）
''''
''''1#      裁减用矩形框
''''2#      裁减用黑色遮盖矩形
''''3#      裁减用黑色遮盖矩形
''''4#      裁减用黑色遮盖矩形
''''5#      裁减用黑色遮盖矩形
''''6#      1#配套的TEXT用
''''7#      体位标注左,四个体位标注的顺序是固定的，一定是左上右下
''''8#      体位标注上
''''9#      体位标注右
''''10#     体位标注下
''''11#     标注句柄（已用）
''''12#     标注句柄（已用）
''''13#     标注句柄（已用）
''''14#     标注句柄（已用）
''''15#     标注句柄（已用）
''''16#     标注句柄（已用）
''''17#     标注句柄（已用）
''''18#     标注句柄（已用）
''''19#     标注句柄（备用）
''''20#     标注句柄（备用）
''''21#     矢冠状重建的竖线
''''22#     矢冠状重建的横线
''''23#     矢冠状重建的边点
''''24#     矢冠状重建的边点
''''25#     矢冠状重建的边点
''''26#     矢冠状重建的边点
''''27#     矢冠状重建的中心点
''''28#     矢冠状重建的结果图轴位线
''''29#     矢冠状重建的结果图矢状位或者冠状位线
''''30#     当前图像窗宽/窗位显示
''''31#     病人四角信息左上
''''32#     病人四角信息左下
''''33#     病人四角信息右下
''''34#     病人四角信息右上
''''35#     病人标尺左
''''36#     病人标尺上
''''37#     病人标尺右
''''38#     病人标尺下
''''39#     病人标尺左的单位
''''40#     病人标尺上的单位
''''41#     病人标尺右的单位
''''42#     病人标尺下的单位
''''43#     打印标记

''''''''''''''''''''''''''''''''''''''''[数据库设置]''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'定义数据库联接

Public Const intSliceOffset = 10


Public cnAccess As ADODB.Connection
Public gcnOracle As ADODB.Connection         ''''ORACLE数据库连接串
Public strADOcn  As String
Public rsTemp As Recordset
Public blLocalRun As Boolean                '观片站是否在本地或单机运行
Public glngUserID  As Long                  '跟“影像界面设置表”默认窗宽窗位等表相关的用户ID
Public PstrCheckUID As String               '最新的检查UID
Public PstrFtpHost As String                'FTP主机地址
Public PstrFtpUser As String                'FTP用户名
Public PstrFtpPwd As String                 'FTP密码
Public PstrFtpPath As String                'FTP目录，是图像文件所在的具体目录
Public PstrBufferImagePath As String        '本机缓存目录
''''''''''''''''''''''''''''''''''''''''权限变量定义'''''''''''''''''''''''''''''''''''''''
Public mstrPrivs As String
Public glngSys As Long                      '系统号
''''''''''''''''''''''''''''''''''''''''[界面布局]'''''''''''''''''''''''''''''''''''''''''
Public intSpaceSize As Integer                          ''序列之间的间隔宽度、高度
Public intMaxAreaX As Long                              ''横向最多可划分的序列个数
Public intMaxAreaY As Long                              ''纵向最多可划分的序列个数
Public Const G_INT_MAX_IMG_COL = 8                      ''横向最多可以划分的图像个数，即图像列数
Public Const G_INT_MAX_IMG_ROW = 8                      ''纵向最多可以划分的图像个数，即图像行数

Public lngDefaultImageBorderColor As Long               ''默认（未选中，非当前）图像边框颜色
Public lngDefaultImageBorderLineStyle As Long           ''默认（未选中，非当前）图像边框线形
Public lngDefaultImageBorderLineWidth As Long           ''默认（未选中，非当前）图像边框线宽

Public lngCurrentImageBorderColor As Long               ''当前图像边框颜色
Public lngCurrentSeriesBorderColor As Long              ''当前（未选中）序列边框颜色
Public lngCurrentImageBorderLineStyle As Long           ''当前图像边框线型
Public lngCurrentImageBorderLineWidth As Long           ''当前图像边框线宽

Public lngSelectedImageBorderColor As Long              ''选中图像边框颜色
Public lngSelectedImageBorderLineStyle As Long          ''选中图像边框线型
Public lngSelectedImageBorderLineWidth As Long          ''选中图像边框线宽

Public lngCellSpacing   As Long                         ''图像间距
Public lngImageIdentifierSize   As Long                 ''图像选择标记大小
Public blnDsipSpilthBorder As Boolean                   ''多余边框是否显示
Public lngSelectImageForeColour As Long                 ''选中图像标识填充色
Public lngViewerBackColor   As Long                     ''Viewer背景颜色
Public lngProgramBackColor As Long                      ''观片工作站底色,程序背景颜色
Public blnDockMiniImage As Boolean                      ''缩略图停靠于菜单下
Public blnShowMiniImageInfo As Boolean                  ''缩略图中显示图像信息
Public blnShowMPRLine As Boolean                        ''MPR操作时，显示位置辅助线
Public blnSquareFrame As Boolean                        ''是否正方形框选报告图
Public blnShowPrintTag As Boolean                       ''是否显示胶片已打印的标记
Public blnPrintFilmBeep As Boolean                      ''胶片打印时是否提示声音，包括添加胶片，打印
Public gblnCompareSize As Boolean                       ''启用FTP文件大小对比

Public Const G_INT_MPR_RADIUS = 16                      ''矢冠状重建的句柄直径

'''''''''''''''''''''''''''''''''''''''''''[LABEL定义]''''''''''''''''''''''''''''''''''''
Public Const G_INT_SYS_LABEL_COUNT = 50                 ''系统内部使用的图像标注总数
Public Const G_INT_SYS_LABEL_HIDE_LEFT = -12000         ''隐藏系统标注时，将标注移动到的左边位置
Public Const G_INT_SYS_LABEL_HIDE_TOP = 12000           ''隐藏系统标注时，将标注移动到的上边位置
Public Const G_INT_SYS_LABEL_TIWEI = 7                  ''体位标注左,上，右，下。四个标注固定顺序相连
Public Const G_INT_SYS_LABEL_WWWL = 30                  ''窗宽窗位标注
Public Const G_INT_SYS_LABEL_PAT_INFO = 31              ''病人四角信息左上，左下，右下，右上。四个标注固定顺序相连，31-34
Public Const G_INT_SYS_LABEL_RULLER = 35                ''病人标尺信息左，上，右，下。四个标注固定顺序相连，后面是四个相连的单位标注，35-42
Public Const G_INT_SYS_LABEL_MPRV = 21                  ''MPR矢冠状位重建的竖线
Public Const G_INT_SYS_LABEL_MPRH = 22                  ''MPR矢冠状位重建的横线
Public Const G_INT_SYS_LABEL_MPR_POINT_V1 = 23          ''MPR矢冠状位重建的竖线第一个端点
Public Const G_INT_SYS_LABEL_MPR_POINT_V2 = 25          ''MPR矢冠状位重建的竖线第二个端点
Public Const G_INT_SYS_LABEL_MPR_POINT_H1 = 24          ''MPR矢冠状位重建的横线第一个端点
Public Const G_INT_SYS_LABEL_MPR_POINT_H2 = 26          ''MPR矢冠状位重建的横线第二个端点
Public Const G_INT_SYS_LABEL_MPR_POINT_O = 27           ''MPR矢冠状位重建的竖线和横线的中心点
Public Const G_INT_SYS_LABEL_MPR_RESULT_H = 28          ''矢冠状重建的结果图轴位投影线
Public Const G_INT_SYS_LABEL_MPR_RESULT_V = 29          ''矢冠状重建的结果图矢状位或者冠状位投影线
Public Const G_INT_SYS_LABEL_PRINT_TAG = 43             '' 打印标记

Public intTextoOffX As Long, intTextoOffY As Long       ''标注文字的偏移量
Public lngLabelColor As Long                            ''标注显示色，白色
Public lngLabelSelectedColor As Long                    ''标注选中色，红色
Public lngLabelLineStyleNorm As Long                    ''标注正常线型
Public lngLabelLineWidthNorm As Long                    ''标注正常线宽
Public lngLabelFontSize As Long                         ''标注文字大小
Public lngLabelLineStyleSel As Long                     ''标注选中线型
Public lngLabelLineWidthSel As Long                     ''标注选中线宽
Public intPeriodSize As Long                            ''选择句柄大小
Public lngPeriodColor As Long                           ''选择句柄颜色
Public blnLabelTextScaleFontSize As Boolean             ''标注文字大小是否随着图像一起缩放
Public intSelectLabelStyle  As Integer                  ''需要画的标注DicomOBJECT类型编码,在按下标注按钮的时候填写
Public bROIArea As Boolean                              ''显示面积
Public bROIMean As Boolean                              ''显示平均值
Public bROIStandardDeviation As Boolean                 ''显示均方差
Public bROILength As Boolean                            ''显示周长
Public bROIMax As Boolean                               ''显示最大值
Public bROIMin As Boolean                               ''显示最小值
Public bROITextChinese As Boolean                       ''测量显示的关联文字信息使用中文
Public lngWinWidthLevelLocation As Long                 ''窗宽窗位的位置 1-上边；2-下边；3-左边；4-右边
Public intNarrowThreshold As Integer                    ''血管狭窄测量的预设阈值
Public intStandardThreshold   As Integer                ''血管狭窄测量的预设阈值
Public intVasEdgeWidth As Integer                       ''血管狭窄测量中显示血管壁短直线的宽度

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public blnRulerDsipLeft As Boolean                      ''是否显示左边标尺
Public blnRulerDsipBottom   As Boolean                  ''是否显示底部标尺
Public blnRulerDsipRight   As Boolean                   ''是否显示右边标尺
Public blnRulerDsipTop   As Boolean                     ''是否显示顶部标尺
Public intRulerWidth As Long                            ''标尺宽度
Public intRulerHeight   As Long                         ''标尺高度
Public intRulerTop   As Long
Public intRulerLeft   As Long
Public lngRulerLeftColor   As Long                      ''标尺颜色
Public intRulerLineWidth   As Long                      ''标尺线宽

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public lngReferenceLineColor   As Long                  ''定位线颜色
Public lngReferenceLineStyle   As Long                  ''定位线线形
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public lngStackStep As Long                             ''穿梭步长
Public lngZoomStep As Long                              ''缩放步长
Public lngCruiseStep As Long                            ''漫游步长
Public lngWidthLevelStep As Long                        ''调窗步长
Public intMouseWheelRoll As Integer                     ''鼠标滚轮滚动的用法
Public intMouseWheelDrag As Integer                     ''鼠标滚轮滚动的用法
''''''''''''''''''''''''''''''''''''''''[病人信息]'''''''''''''''''''''''''''''''''''''''''
Public blnAnatomicMarkersLeft As Boolean                ''是否显示左边体位标记
Public blnAnatomicMarkersTop   As Boolean               ''是否显示顶部体位标记
Public blnAnatomicMarkersBottom   As Boolean            ''是否显示底部体位标记
Public blnAnatomicMarkersRight   As Boolean             ''是否显示右边体位标记
Public blnChinaMark   As Boolean                        ''是否采用汉字显示体位标记
Public lngPatientInfoInvisibleSize As Long              ''图像小于X时，不显示病人信息
Public lngpatientInfoColor As Long                      ''病人信息颜色
Public blnpatientInfoScaleFontSize As Boolean           ''病人信息文字大小是否随着图像一起缩放
Public blnHidePatientInfo As Boolean                    ''是否显示病人信息

''''''''''''''''''''''''''''''''''''''''[图像插值]'''''''''''''''''''''''''''''''''''''''''
Public Const intMagnificationMode = 3                   ''图像插值模式

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public lngPatientInfoFontSize As Long                   ''病人信息显示字体大小
Public blnPatientInfoFontBold As Boolean                ''病人信息显示字体粗体
Public blnPatientInfoFontItalic As Boolean              ''病人信息显示字体斜体
Public strPatientInfoFontName As String                 ''病人信息显示字体名称
Public lngPatientInfoTitle As Long                      ''病人信息使用的题头，0--不使用题头；1--中文题头；2--英文题头
Public bShowFilmConfig As Boolean                       ''在点击照相按钮时，是否弹出胶片设置窗口

Public blnInterfaceParaModified As Boolean              ''记录影像界面系统参数的值是否发生改变？

''''''''''''''''''''''''''''''''''''''''''''''''''''

Public lngReferenceLineSpacing As Long                  ''定位线的显示间距
Public cstrPrintAE As String                            ''打印时使用的本程序AE名称
Public intFilmFontSize As Integer                       ''记录打印胶片使用的标注文字大小
Public blnPrintOkEcho As Boolean                        ''打印完成后，弹出提示对话框

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''[菜单控制的临时变量]''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Button_miSerialPlaceInPhase  As Boolean           ''序列同步
Public Button_miSerialManualSyn As Boolean               ''手工序列同步
Public Button_miImageInPhase As Boolean                  ''图像状态同步
Public Button_miLookOrBrowse As Boolean                  ''浏览观察模式
Public Button_miCutOut As Boolean                        ''裁剪
Public Button_miFrameSelectImage As Boolean              ''框选图象
Public Button_miStack As Boolean                         ''穿梭
Public Button_miWidthLevel As Boolean                    ''手动调窗
Public Button_miZoom As Boolean                          ''缩放
Public Button_miCruise As Boolean                        ''漫游
Public Button_mi3dCursor As Boolean                      ''3D鼠标
Public Button_miAutoWidthLevel As Boolean                ''自适应调窗
Public Button_miDispPatientInfo As Boolean               ''属性显示
Public Button_miLabeltext As Boolean                     ''文字
Public Button_miDispLabelInfo As Boolean                 ''标注显示
Public Button_miLabelAngle As Boolean                    ''角度
Public Button_miLabelPolygon As Boolean                  ''区域
Public Button_miAllReferLine As Boolean                  ''所有定位线
Public Button_miFLReferLine As Boolean                   ''首尾定位线
Public Button_miCurrentReferLine As Boolean              ''当前定位线
Public Button_miLabelRectangle As Boolean                ''矩形
Public Button_miLabelLine As Boolean                     ''直线
Public Button_miLabelEllipse As Boolean                  ''椭圆
Public Button_miLabelArrowhead As Boolean                ''箭头
Public Button_miLabelPolyLine As Boolean                 ''曲线
Public Button_miLabelVasMeasure As Boolean               ''血管狭窄测量
Public Button_miLabelCadiothoracicRatio As Boolean       ''心胸比测量
Public Button_miFullScreen As Boolean                    ''全屏显示
Public Button_miMouseShowValue As Boolean                ''在鼠标上显示CT值
Public Button_miShowMiniSeries As Boolean                ''显示序列缩略图
Public Button_miViewAllSeries As Boolean                 ''全序列观片
Public Button_miShowOverlay As Boolean                   ''显示Overlay


'''''''''''''''''''''''''菜单和按钮常量定义'''''''''''''''''''''''''''''''''''''''
'---------------------------------------------
'-----------------文件菜单--------------------
'---------------------------------------------
'预留到100现已在使用到15
Public Const ID_File = 101                                               ''文件
Public Const ID_File_Open = 102                                          ''打开文件
Public Const ID_File_Close = 103                                         ''关闭序列
Public Const ID_File_DelAllPhoto = 104                                   ''删除所有图像
Public Const ID_File_DelReport = 105                                     ''删除报告图像
Public Const ID_File_SaveFile = 106                                      ''保存文件
Public Const ID_File_SaveASFile = 107                                    ''另存文件
Public Const ID_File_SaveToCD = 115                                      ''创建CD
Public Const ID_File_SAveASReport = 108                                  ''保存报告图
'***************************************
Public Const ID_File_Send = 109                                          ''发送
Public Const ID_File_Send_GetHost = 110                                  ''接收主机
Public Const ID_File_Send_OutPowerPoint = 111                            ''输出到PowerPoint
'***************************************
Public Const ID_File_OpenDicomDir = 114                                  ''打开DICOMDIR
Public Const ID_File_PhotoProperty = 112                                 ''图像属性
Public Const ID_File_Exit = 113                                          ''退出
'-----------------------------------------------
'------------------视图菜单---------------------
'-----------------------------------------------
'预留到200-300现在已使用251
Public Const ID_View = 200                                              ''视图
Public Const ID_View_Typeset = 201                                      ''版面安排
Public Const ID_View_OneBrowse = 202                                    ''单序列观察
Public Const ID_View_PropertyShow = 203                                 ''属性显示
Public Const ID_View_LableShow = 204                                    ''标注显示
Public Const ID_View_UpSeries = 247                                     ''上一序列
Public Const ID_View_DownSeries = 248                                   ''下一序列
Public Const ID_View_ShowMiniSeries = 249                               ''显示序列缩略图
Public Const ID_View_ViewAllSeries = 250                                ''全序列观片
Public Const ID_View_ShowOverlay = 251                                  ''显示Overlay
'***********************************************
Public Const ID_View_PhotoSerial = 205                                  ''图像顺序
Public Const ID_View_PhotoSerial_PhotoNumber = 206                      ''图像号
Public Const ID_View_PhotoSerial_BedASC = 207                           ''床位正序
Public Const ID_View_PhotoSerial_BedDESC = 208                          ''床位逆序
Public Const ID_View_PhotoSerial_CollectionTime = 209                   ''采集时间
Public Const ID_View_PhotoSerial_PhotoTime = 210                        ''图像时间
'************************************************
Public Const ID_View_ShowScale = 230                                    ''显示比例
Public Const ID_View_ShowScale_AutoShow = 240                           ''自适应
Public Const ID_View_ShowScale_50% = 241                                ''50%
Public Const ID_View_ShowScale_100% = 242                               ''100%
Public Const ID_View_ShowScale_200% = 243                               ''200%
Public Const ID_View_showScale_400% = 244                               ''400%
Public Const ID_View_showScale_150% = 2411                              ''150%
Public Const ID_View_showScale_250% = 2412                              ''250%
Public Const ID_View_showScale_300% = 2413                              ''300%
Public Const ID_View_ShowScale_Custom = 245                             ''自定义
'************************************************
Public Const ID_View_FullScreen = 246                                   ''全屏显示
''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''动作菜单''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''
'预留到300-400现在已使用367,其中349-360是快捷键调窗的ID
Public Const ID_Active = 300                                            ''动作
'***********************************************
Public Const ID_Active_Select = 301                                     ''选择
Public Const ID_Active_Select_OneSelect = 302                           ''单幅选择
Public Const ID_Active_Select_SelectAllSerial = 303                     ''选择所有序列
Public Const ID_Acitve_Select_SelectAllPhoto = 304                      ''选择所有图像
'************************************************
Public Const ID_Active_Also = 305                                       ''同步
Public Const ID_Active_Also_Serial = 306                                ''序列同步
Public Const ID_Active_Also_Photo = 307                                 ''图像同步
Public Const ID_Active_Also_ManualSerial = 363                          ''手工序列同步
Public Const ID_Active_Also_LockSerial = 364                            ''锁定序列
'************************************************
Public Const ID_Active_Shuttle = 308                                    ''穿梭
Public Const ID_Active_Cruise = 309                                     ''漫游
Public Const ID_Active_Cut = 310                                        ''裁剪
Public Const ID_Active_Zoom = 311                                       ''缩放
Public Const ID_Active_ReSetAll = 312                                   ''恢复所有
'************************************************
Public Const ID_Active_AdjustWindow = 313                               ''调窗
Public Const ID_Active_AdjustWindow_HandAdjustWindow = 314              ''手控调窗
Public Const ID_Active_AdjustWindow_HandAdjustWindow_ReSet = 349        ''手控调窗_恢复
Public Const ID_Active_AdjustWindow_HandAdjustWindow_Custom = 350       ''手控调窗_自定义
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F3 = 351           ''手控调窗_F3
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F4 = 352           ''手控调窗_F4
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F5 = 353           ''手控调窗_F5
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F6 = 354           ''手控调窗_F6
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F7 = 355           ''手控调窗_F7
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F8 = 356           ''手控调窗_F8
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F9 = 357           ''手控调窗_F9
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F10 = 358          ''手控调窗_F10
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F11 = 359          ''手控调窗_F11
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F12 = 360          ''手控调窗_F12
Public Const ID_Active_AdjustWindow_AutoAdjustWindow = 315              ''自适应调窗
Public Const ID_Active_AdjustWidnow_CustomAdjustWindow = 316            ''自定义调窗
'************************************************
Public Const ID_Active_PointingLine = 317                               ''定位线
Public Const ID_Active_PointingLine_ALL = 318                           ''所有定位线
Public Const ID_Active_PointingLine_FirstLast = 319                     ''首位定位线
Public Const ID_Active_PointingLine_Now = 320                           ''当前定位线
Public Const ID_Active_PointingLine_3DLine = 321                        ''3D鼠标定位
'************************************************
Public Const ID_Active_Eddy = 322                                       ''旋转
Public Const ID_Active_Eddy_LeftRight = 323                             ''左右翻转
Public Const ID_Active_Eddy_TopButton = 324                             ''垂直翻转
Public Const ID_Active_Eddy_Left90 = 325                                ''左旋90
Public Const ID_Active_Eddy_Right90 = 326                               ''右旋90
'************************************************
Public Const ID_Active_ReverseVideo = 327                               ''反白
'************************************************
Public Const ID_Active_SieveLens = 328                                  ''滤镜
Public Const ID_Active_SieveLens_Model = 32810                          ''常用滤镜模板，从32810开始到32850，最多支持40个
Public Const ID_Active_SieveLens_LancetMinus = 329                      ''边缘增强强度减少
Public Const ID_Active_SieveLens_LancetAdd = 330                        ''边缘增强强度增加
Public Const ID_Active_SieveLens_FlatnessMinus = 331                    ''平滑减少
Public Const ID_Active_SieveLens_FlatnessAdd = 332                      ''平滑增加
Public Const ID_Active_Sievelens_LeftMoveMinus = 333                    ''边缘增强幅度减少
Public Const ID_Active_Sievelens_LeftMoveAdd = 334                      ''边缘增强幅度增加
Public Const ID_Active_Sievelens_PhotoReset = 335                       ''图像还原
'************************************************
Public Const ID_Active_Lable = 336                                      ''标注
Public Const ID_Active_Lable_Text = 337                                 ''文字
Public Const ID_Active_Lable_Arrowhead = 338                            ''箭头
Public Const ID_Active_Lable_Ellipse = 339                              ''椭圆
Public Const ID_Active_Lable_Angle = 340                                ''角度
Public Const ID_Active_Lable_Curve = 341                                ''曲线
Public Const ID_Active_Lable_Area = 342                                 ''区域
Public Const ID_Active_Lable_BeeLine = 343                              ''直线
Public Const ID_Active_Lable_Rect = 344                                 ''矩形
Public Const ID_Active_Lable_AreaBeeLinePhoto = 345                     ''区域直方图
Public Const ID_Active_Lable_AdjustLine = 346                           ''校准
Public Const ID_Active_Lable_ClearLbale = 347                           ''清除标注
Public Const ID_Active_Lable_DelSelectLable = 348                       ''删除标注
Public Const ID_Active_Lable_VasMeasure = 361                           ''狭窄血管测量
Public Const ID_ACtive_Mouse_Value = 362                                ''在鼠标上显示CT值
Public Const ID_ACtive_FrameSelectImage = 365                           ''框选图象
Public Const ID_ACtive_SaveInReport = 366                               ''当前图像存成报告图
Public Const ID_Active_Lable_CadioThoracicRatio = 367                   ''心胸比测量
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''工具菜单''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'预留到400-500现在已使用420
Public Const ID_Tool = 400                                              ''工具
Public Const ID_Tool_Movie = 401                                        ''电影
Public Const ID_Tool_Magnifier = 402                                    ''放大镜
Public Const ID_Tool_ArrowyCoronaryReset = 403                          ''矢冠状重建
Public Const ID_Tool_NumberMinusShadow = 404                            ''数字减影
Public Const ID_Tool_BogusColour = 405                                  ''伪彩观察
Public Const ID_Tool_FilmPrint = 406                                    ''胶片打印
Public Const ID_Tool_Film_AddSeries = 40601                             ''胶片打印--打印当前序列
Public Const ID_Tool_Film_AddImage = 40602                              ''胶片打印 -- 打印当前图像
Public Const ID_Tool_Film_AddSelected = 40603                           ''胶片打印 -- 打印当前选择
Public Const ID_Tool_Film_AddInterval = 40604                           ''胶片打印 -- 间隔打印当前序列
Public Const ID_Tool_PhotoUnite = 407                                   ''图像拼接
Public Const ID_Tool_LableTool = 408                                    ''标注工具
Public Const ID_Tool_LookPhotoOption = 409                              ''观片选项
'*************************************
Public Const ID_ToolBar = 410                                           ''工具栏
Public Const ID_ToolBar_Left = 411                                      ''靠左
Public Const ID_ToolBar_Right = 412                                     ''靠右
Public Const ID_ToolBar_Top = 413                                       ''靠上
Public Const ID_ToolBar_Button = 414                                    ''靠下
Public Const ID_toolBar_16Icon = 415                                    ''16*16图标
Public Const ID_ToolBar_24Icon = 416                                    ''24*24图标
Public Const ID_ToolBar_32Icon = 417                                    ''32*32图标
Public Const ID_ToolBar_Hide = 418                                      ''隐藏工具栏
Public Const ID_Tool_NothinMouseState = 419                             ''清除所有鼠标使用状态
Public Const ID_Tool_SlopeReconstruction = 420                          ''斜面重建
'*************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''帮助菜单''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''弹出菜单'''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''预留到500---800'''''''''''''''''''''
'''''''''增加frmImageSpelling窗体工具按钮定义801-807'''''''''
Public Const ID_frmImageSpelling_CompleteSpelling = 801                 ''完成拼接
Public Const ID_frmImageSpelling_SavePhoto = 802                        ''保存图像
Public Const ID_frmImageSpelling_DelPhoto = 803                         ''删除图像
Public Const ID_frmImageSpelling_ZoomOut = 804                          ''缩放图像
Public Const ID_frmImageSpelling_Quit = 806                             ''退出
Public Const ID_frmImageSpelling_CutOut = 807                           ''裁剪图像
Public Const ID_frmImageSpelling_Move = 808                             ''移动图像
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''增加frmFilm窗体工具按钮定义830-872'''''''''''''''
Public Const ID_frmFilm_TakePictures = 830                              ''照相
Public Const ID_frmFilm_FilmCol = 831                                   ''纵向
Public Const ID_frmFilm_FilmRow = 832                                   ''横向
Public Const ID_frmFilm_RectPhotCase = 833                              ''正方形图像格
Public Const ID_frmFilm_FormatCustom = 834                              ''格式定义
Public Const ID_frmFilm_FilmSize = 835                                  ''胶片大小
Public Const ID_frmFilm_Format = 836                                    ''格式
Public Const ID_frmFilm_Camera = 837                                    ''相机
Public Const ID_frmFilm_Quit = 838                                      ''退出
Public Const ID_frmFilm_DeleteImg = 839                                 ''删除图像
Public Const ID_frmFilm_WinLevel = 840                                  ''调窗
Public Const ID_frmFilm_Pan = 841                                       ''漫游
Public Const ID_frmFilm_Zoom = 842                                      ''缩放
Public Const ID_frmFilm_RotateLeft = 843                                ''向左旋转
Public Const ID_frmFilm_RotateRight = 844                               ''向右旋转
Public Const ID_frmFilm_FlipHorizontal = 845                            ''左右镜象
Public Const ID_frmFilm_FlipVertical = 846                              ''上下镜象
Public Const ID_frmFilm_Resume = 847                                    ''恢复
Public Const ID_frmFilm_ImgSynchronal = 848                             ''图像同步
Public Const ID_frmFilm_Divide = 851                                    ''图像分格
Public Const ID_frmFilm_UnDivide = 852                                  ''取消分格
Public Const ID_frmFilm_Invert = 853                                    ''反白
Public Const ID_frmFilm_SelAll = 854                                    ''全选
Public Const ID_frmFilm_RectZoom = 855                                  ''框选缩放
Public Const ID_frmFilm_CutOut = 856                                    ''裁剪
Public Const ID_frmFilm_CutOut_14X17 = 85601                            ''裁剪，固定比例，14*17
Public Const ID_frmFilm_CutOut_11X14 = 85602                            ''裁剪，固定比例，11*14
Public Const ID_frmFilm_CutOut_10X14 = 85603                            ''裁剪，固定比例，10*14
Public Const ID_frmFilm_CutOut_8X10 = 85604                             ''裁剪，固定比例，8*10
Public Const ID_frmFilm_CutOut_14X14 = 85605                            ''裁剪，固定比例，14*14
Public Const ID_frmFilm_CutOut_17X14 = 85606                            ''裁剪，固定比例，17*14
Public Const ID_frmFilm_CutOut_14X11 = 85607                            ''裁剪，固定比例，14*11
Public Const ID_frmFilm_CutOut_14X10 = 85608                            ''裁剪，固定比例，14*10
Public Const ID_frmFilm_CutOut_10X8 = 85609                             ''裁剪，固定比例，10*8
Public Const ID_frmFilm_CutOut_Custom = 85610                           ''裁剪，自由比例

Public Const ID_frmFilm_FilterLengthUp = 857                            ''平滑增加
Public Const ID_frmFilm_FilterLengthDown = 858                          ''平滑减少
Public Const ID_frmFilm_OpenImages = 859                                ''打开图象
Public Const ID_frmFilm_Label = 860                                     ''标注文字下拉菜单
Public Const ID_frmFilm_Label_A = 861                                   ''标注文字-Anterior-前
Public Const ID_frmFilm_Label_P = 862                                   ''标注文字-Posterior-后
Public Const ID_frmFilm_Label_L = 863                                   ''标注文字-Left-左
Public Const ID_frmFilm_Label_R = 864                                   ''标注文字-Right-右
Public Const ID_frmFilm_Label_S = 865                                   ''标注文字-Superior-上
Public Const ID_frmFilm_Label_I = 866                                   ''标注文字-Inferior-下
Public Const ID_frmFilm_Label_Delete = 867                              ''删除标志文字
Public Const ID_frmFilm_SelNone = 868                                   ''选择--全清图像
Public Const ID_frmFilm_SelSeries = 869                                 ''选择 -- 选择序列
Public Const ID_frmFilm_SelInverse = 870                                ''选择 - 反选
Public Const ID_frmFilm_ImgIncrease = 871                               ''排序 - 正序
Public Const ID_frmFilm_ImgDecrease = 872                               ''排序 - 逆序

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''增加frmPacsImg窗体工具按钮定义880-884'''''''''''''''
Public Const ID_PacsImg_SelectAllSeries = 880                           ''全选序列
Public Const ID_PacsImg_UnSelectAllSeries = 881                         ''全清序列
Public Const ID_PacsImg_SelectAllImages = 882                           ''全选图像
Public Const ID_PacsImg_UnSelectAllImages = 883                         ''全清图像
Public Const ID_PacsImg_ReverseSelectImages = 884                       ''反选图象


Public Const ID_Help = 600                                              ''帮助
Public Const ID_Help_Help = 601                                         ''帮助
Public Const ID_Help_WebZLSOFT = 602                                    ''WEB上的中联
Public Const ID_Help_WebZLSOFT_WEB = 603                                ''中联主页
Public Const ID_Help_WebZLSOFT_Mail = 604                               ''发送反馈
Public Const ID_Help_About = 605                                        ''关于
Public Const ID_Help_UpdateDB = 606                                     ''升级数据库
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''快捷按钮''''''''''''''''''''''''''''''''''''
Public Const FSHIFT = 4                                            'Shift
Public Const FCONTROL = 8                                          'Ctrl
Public Const FALT = 16                                             'ALT
Public Const VK_F1 = &H70                                          'F1
Public Const VK_F2 = &H71                                          'F2
Public Const VK_F3 = &H72                                          'F3
Public Const VK_F4 = &H73                                          'F4
Public Const VK_F5 = &H74                                          'F5
Public Const VK_F6 = &H75                                          'F6
Public Const VK_F7 = &H76                                          'F7
Public Const VK_F8 = &H77                                          'F8
Public Const VK_F9 = &H78                                          'F9
Public Const VK_F10 = &H79                                         'F10
Public Const VK_F11 = &H7A                                         'F11
Public Const VK_F12 = &H7B                                         'F12
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''工具条常量''''''''''''''''''''''''''''''''''''
Public Const ToolBar_Menu  As Integer = 1                               ''菜单
Public Const ToolBar_Main  As Integer = 2                               ''主工具条
Public Const ToolBar_Photo  As Integer = 3                              ''图像处理工具条
Public Const ToolBar_Scale As Integer = 4                               ''测量工具条
Public Const ToolBar_Plane  As Integer = 5                              ''平面工具条
Public Const ToolBar_Object  As Integer = 6                             ''对像工具条
Public Const ToolBar_Comm  As Integer = 7                               ''公共工具条
Public Const toolBar_PhotoStrong As Integer = 8
'''''''''''''''''''''''当前工具条设置''''''''''''''''''''''''''''''''
Public intToolBarIconSize As Integer                                    ''图标大小
Public intToolBarPosition As Integer                                    ''摆放位置
Public blToolBarHide As Boolean                                         ''隐藏工具条

''''''''''''''''''''''工具条风格'''''''''''''''''''''''''''''''''''''
Public IntComBarTheme As Integer                                        ''统一工具条的显示风格
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''[系统参数设置]''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public cLabelStore As New Collection            '记录保存标注所使用的图像头信息
Public Const cProducer = "ZLPACS"
Public intStatusBarFontSize As Integer

'''''''''''''''''''''''''''[预设窗宽窗位 F3---F12]''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type TPresetWinWL
    bInUse As Boolean                               ''标识本快捷键是否被设置
    strModality As String                           ''影像类别
    strWinWLCName As String                         ''快捷键上设置的窗宽窗位中文名称
    strWinWLEName As String                         ''快捷键上设置的窗宽窗位英文名称
    lngWinWidth As Long                             ''快捷键上设置的窗宽窗值
    lngWinLevel As Long                             ''快捷键上设置的窗位值
    intDefault As Integer                           ''是否默认窗宽窗位
    lngID As Long                                   ''预设窗宽窗位的ID
End Type
Public aPresetWinWL() As TPresetWinWL         ''保存预设窗宽窗位的数组，
                                                    ''允许的快捷键值为F3--F12，对应于数组的下标

'''''''''''''''''''''''''''[预设屏幕布局]''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type TModifiedPresetLayout
    bModified As Boolean
    strModality As String                               ''记录对应的影像类别
    bSeriesAutoFormat As Boolean                        ''打开窗体时是否自动安排序列格式
    lngSeriesRows As Long                               ''预设打开窗体时使用的序列行数
    lngSeriesColumns As Long                            ''预设打开窗体时使用的序列列数
    bImageAutoFormat As Boolean                         ''打开图像时是否自动安排图像格式
    lngImageRows As Long                                ''预设打开图像时使用的图像行数
    lngImageColumns As Long                             ''预设打开图像时使用的图像列数
    bInvert As Boolean                                  ''打开图像时是否自动反白，【无用】
    bShowPatientInfo As Boolean                         ''打开图像时是否显示病人信息,【无用】
    bAutoSelectReferenceLine As Boolean                 ''打开图像时是否自动选择显示定位线（只针对CT,MR图像有此设置）,【无用】
    bAutoSelectSeriesSyn As Boolean                     ''打开图像时是否自动选择序列间图像位置同步（只针对CT,MR图像有此设置）,【无用】
    lngInterpolationMode As Long                        ''图像放大时的插值模式
    lngImageSort As Long                                ''图像排序方式：0-默认；1-图像号；2-床位正序；3-床位逆序；4-采集时间；5-图像时间
End Type

Public aPresetLayout() As TModifiedPresetLayout         ''保存预设屏幕布局的数组
Public aModifiedPresetLayout() As TModifiedPresetLayout ''保存被修改的屏幕布局的数组

'''''''''''''''''''''''''''''''''''''[预设图像消隐]'''''''''''''''''''''''''''''''''''''''''''''''''
Public Type TImageShutter
    bModified As Boolean            ''是否被修改
    strModality As String           ''记录对应的影像类别
    intShutterType As Integer       ''消隐的类型：0－无消隐；1－圆形消隐；2－矩形消隐；4－多边形消隐。
                '这些消隐类型可以相互叠加，但是每种类型只能够被叠加一次。例如，同时使用圆形和多边形消隐，
                '则消隐类型为1+4＝5。消隐类型大于7则认为无效类型，自动设置为0。
    intCenterX As Integer           ''圆形消隐的圆心X坐标
    intCenterY As Integer           ''圆形消隐的圆心Y坐标
    intRadius As Integer            ''圆形消隐的半径
    intRectLeft As Integer          ''矩形消隐的左边界
    intRectRight As Integer         ''矩形消隐的右边界
    intRectUpper As Integer         ''矩形消隐的上边界
    intRectLower As Integer         ''矩形消隐的下边界
    strVertices As String           ''多边形消隐的顶点集，使用英文字符的“：”来间隔。
    lngColor As Long                ''消隐的灰度颜色
End Type

Public aImageShutter() As TImageShutter             ''保存预设的图像消隐设置
Public aModifiedImageShutter() As TImageShutter     ''保存被修改的图像消隐设置

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''[鼠标用法设置]'''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public iMouseFuncCount As Integer                       ''鼠标功能键的数量
Public cMouseUsage As New Collection                  ''记录鼠标用法的集合
Public cModifiedMouseUsage As New Collection          ''临时记录鼠标用法被修改状态的集合
Public bMouseUsageModified As Boolean
Public Const lngDrawLabelFuncNo = 20                    ''画标注的综合功能序号
Public Const lngDrawLabelCurrent = 1                    ''画标注功能被选择为当前鼠标按钮后，真正被选的按钮功能号
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''[病人信息标注位置和显示设置]''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type TInfoLabelLocate                            ''标识图像信息在图像上四个角的显示位置
    lngID As Long                                       ''在数据库中的ID号
    strGroup As String                                  ''图像内dicom标识的Group号
    strElement As String                                ''图像内dicom标识的Element号
    strEName As String                                  ''图像信息的英文名
    strCName As String                                  ''图像信息的中文名
    bUsed As Boolean                                    ''该信息是否被选用
    lngLocation As Long                                 ''标识信息所在的位置
    lngOrder As Long                                    ''表示信息在被选中角内的位置序号
    blnIsExport As Boolean                              ''表示是否允许导出此信息
End Type
Public aInfoLabelLocate() As TInfoLabelLocate           ''保存图像信息显示方式的数组
Public lngInfoLabelCount As Long                        ''记录可以使用的图像信息数量
Public bInfoLabelModified As Boolean                    ''记录病人四角信息的设置是否被改变了

Public cDICOMPrinter As New Collection                  ''[DICOM打印机参数设置]
Public blnSelectedImageIfColor As Boolean               ''当前选图像是否是彩色图像

'''''''''''''''''''''''''''''''''''[胶片打印，图像布局]''''''''''''''''''''''''

Public Const CUT_LABEL = "CUT"
Public Const POSTURE_LABEL = "体位"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''千图读片'''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public ZLSeriesInfos As Collection          ''记录当前打开的所有图像文件信息
Public ZLShowSeriesInfos As Collection    '记录已经显示的序列中图像的状态

Public Const ATTR_影像类别 As String = "8:60"
Public Const ATTR_序列号 As String = "20:11"
Public Const ATTR_图像号 As String = "20:13"

Public Const ATTR_采集日期 As String = "8:22"
Public Const ATTR_采集时间 As String = "8:32"
Public Const ATTR_图像日期 As String = "8:23"
Public Const ATTR_图像时间 As String = "8:33"
Public Const ATTR_层厚 As String = "18:50"
Public Const ATTR_图像位置病人 As String = "20:32"
Public Const ATTR_图像方向病人 As String = "20:37"
Public Const ATTR_参考帧UID As String = "20:52"
Public Const ATTR_切片位置 As String = "20:1041"
Public Const ATTR_行数 As String = "28:10"
Public Const ATTR_列数 As String = "28:11"
Public Const ATTR_像素距离 As String = "28:30"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''MPR'''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Type MPRCube
    ZLShowSeriesInfos As clsSeriesInfo  '--- 原有的ZLShowSeriesInfos结构
    Images As New DicomImages           '--- 原来Viewer中已经加载的图像，方便恢复图象中的标注、调窗、缩放等信息
    blnIsMPR As Boolean                 '--- 是否当前做MPR的序列，如果是，恢复的时候，不需要替换该序列的内容。
    intViewerIndex As Integer           '--- 摆放重建结果的Viewer的Index
End Type
Public ZLMPRCube(1 To 3) As MPRCube
Public ZLMPRSeriesUID As String         '当前三维重建的序列UID
Public ZLMPRSlopeSeriesUID As String    '当前MPR斜面重建的序列UID


'---------------------------观片站和胶片打印机数量控制，注册-------------------------------
Public Const LOGIN_TYPE_医技观片站 As String = "影像观片站数量"
Public Const LOGIN_TYPE_胶片打印机 As String = "影像胶片打印机数量"
Public gint医技观片站数量 As Integer
Public gint胶片打印机 As Integer



'''''''''''''''''''''''''''[预设滤镜操作]''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type TPresetFilter
    lngID As Long                                   ''预设滤镜的ID
    strname As String                               ''预设滤镜的名称
    strModality As String                           ''影像类别
    intUnSharpEnhancementUp As Integer              ''图像增强强度增加
    intUnSharpEnhancementDown   As Integer          ''图像增强强度减少
    intUnSharpLengthUp  As Integer                  ''图像增强幅度增加
    intUnSharpLengthDown    As Integer              ''图像增强幅度减少
    intFilterLengthUp As Integer                    ''图像平滑增加
    intFilterLengthDown As Integer                  ''图像平滑减少
End Type
Public aPresetFilter() As TPresetFilter             ''保存预设滤镜的数组，
                                        

'''''''''''''''''''''''''''''[图像同步]'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const IMG_SYN_All = 0                            ''全部同步
Public Const IMG_SYN_WINDOW = 1                         ''调窗同步
Public Const IMG_SYN_ZOOMPAN = 2                        ''缩放、漫游同步
Public Const IMG_SYN_ROTATE = 3                         ''旋转同步
Public Const IMG_SYN_FLIP = 4                           ''镜像同步
Public Const IMG_SYN_FILTER = 5                         ''滤镜同步

