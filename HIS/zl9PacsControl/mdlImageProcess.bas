Attribute VB_Name = "mdlImageProcess"
Option Explicit

Public Enum TImageType
    mtTagImage = 0      '标记图
    mtReportImage = 1   '报告图
    mtStadyImage = 2    '检查图
End Enum

Public gobjImageProcess As frmImageProcess

Public glngColor(10) As Long             '标记图中圆形编号使用的9个颜色

Public Const G_STR_TAG = "Po=息肉[+]E=糜烂区[+]M=镶嵌[+]L=粘膜白斑[+]C=湿疣[+]I=浸润性癌[+]W=醋酸白色上皮[+]AT=异常转化区[+]V=非典型血管[+]P=点状血管[+]Xn=直接活检部位"

'图像处理
Public Const conMenu_Process_Window = 501           '亮度对比度
Public Const conMenu_Process_Zoom = 502             '缩放
Public Const conMenu_Process_Corp = 512             '拖动
Public Const conMenu_Process_RRotate = 503          '顺时针旋转
Public Const conMenu_Process_LRotate = 504          '逆时针旋转
Public Const conMenu_Process_Sharpness = 505        '锐化
Public Const conMenu_Process_Filter = 506           '平滑
Public Const conMenu_Process_Arrow = 507            '箭头标注
Public Const conMenu_Process_Ellipse = 508          '圆形标注
Public Const conMenu_Process_Text = 509             '文字标注
Public Const conMenu_Process_RectZoom = 510         '裁剪采集
Public Const conMenu_Process_RectCapture = 511      '裁剪后采集
Public Const conMenu_Process_Line = 520             '直线标注
Public Const conMenu_Process_Exit = 2613            '退出
Public Const conMenu_Process_Save = 3091            '保存
Public Const conMenu_Process_SaveToReport = 3941    '保存到检查
Public Const conMenu_Process_SaveToStady = 3943     '保存到报告
Public Const conMenu_Process_DelAllLabels = 8113    '删除全部标注，使用其他系统的图标编号
Public Const conMenu_Process_MoveLabel = 6891       '移动或删除选中标注，使用其他系统的图标编号
Public Const conMenu_Process_LabelSetUp = 10003     '标注按钮设置，使用其他系统的图标编号
Public Const conMenu_Process_Restore = 8124         '恢复
Public Const conMenu_Process_TextTag = 5010         '文本标记
Public Const conMenu_Process_NumTag = 7405          '数字标记
Public Const conMenu_Process_Page = 1001
Public Const conMenu_Process_Num = 96
Public Const conMenu_Process_Word = 97
