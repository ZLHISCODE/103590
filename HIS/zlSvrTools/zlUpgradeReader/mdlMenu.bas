Attribute VB_Name = "mdlMenu"
Option Explicit

'********************************************************************
'CommandBar命令ID
Public Enum CommandBarIDCond
    conMenu_File = 1
    conMenu_Edit = 2
    conMenu_View = 8
    conMenu_Help = 9
    
    '文件菜单
    conMenu_File_Open = 101
    conMenu_File_Save = 103
    conMenu_File_Login = 107
    conMenu_File_Logout = 108
    conMenu_File_Exit = 109
    
    '编辑菜单

    
    '查看菜单
    conMenu_View_Expend = 711               '展开/折叠组(&X)
    conMenu_View_Expend_AllCollapse = 7111     '折叠所有组(&L)
    conMenu_View_Expend_AllExpend = 7112       '展开所有组(&X)
    conMenu_View_Expend_CurCollapse = 7113     '折叠当前组(&C)
    conMenu_View_Expend_CurExpend = 7114       '展开当前组(&E)

    conMenu_View_ShowPrivewText = 722          '显示需求
    conMenu_View_ShowGroupBox = 723            '显示分组
    conMenu_View_ShowRelation = 724            '显示关联问题
    
    conMenu_View_Filter = 802
    conMenu_View_RecordPrev = 803
    conMenu_View_RecordNext = 804
    conMenu_View_Find = 805
    conMenu_View_FindNext = 806
    conMenu_View_Refresh = 809
    conMenu_View_Close = 810
    
    '帮助菜单
    conMenu_Help_About = 901
    
    conMenu_Custom_System = 900001
    conMenu_Custom_Icon = 900002
End Enum

'CommandBar固有常量定义
Public Const XTP_ID_WINDOW_LIST = 35000 '窗体列表
Public Const XTP_ID_TOOLBARLIST = 59392 '工具栏列表
Public Const ID_INDICATOR_CAPS = 59137 '状态栏（大写）
Public Const ID_INDICATOR_NUM = 59138 '状态栏（数字）
Public Const ID_INDICATOR_SCRL = 59139 '状态栏（滚动）

'CommandBar辅助热键
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16
'********************************************************************

Public Enum IconID
    ICON_Mail = 0   '邮件图标
    ICON_Importance '重要性
    ICON_FlagTrain  '培训标志
    ICON_NoRead     '未读
    ICON_Read       '已读
    ICON_Unknown    '不确定
    ICON_Low        '低
    ICON_Center     '中
    ICON_High       '高
    ICON_Train      '已培训
    ICON_UnTrain    '未培训
End Enum
