Attribute VB_Name = "mdlPubDefine"
Option Explicit


'排队叫号系统
Public Const conMenu_Queue_CallThis = 8250      '直呼
Public Const conMenu_Queue_CallNext = 8251      '顺呼，呼叫下一个

Public Const conMenu_Queue_Restore = 8253       '恢复
Public Const conMenu_Queue_ReCall = 8254        '重呼
Public Const conMenu_Queue_Abandon = 8255       '弃号
Public Const conMenu_Queue_Refresh = 8256       '刷新
Public Const conMenu_Queue_Setup = 8257         '参数设置
Public Const conMenu_Queue_Update = 8258        '修改
Public Const conMenu_Queue_Broadcast = 8259     '广播
Public Const conMenu_Queue_Pause = 8260         '暂停
Public Const conMenu_Queue_Finaled = 8261       '完成就诊
Public Const conMenu_Queue_Find = 8262          '查找
Public Const conMenu_Queue_ComeBack = 8263      '回诊
Public Const conMenu_Queue_RecDiagnose = 8264   '接诊

Public Const conMenu_Queue_Locate = 8265        '定位
Public Const conMenu_Queue_LocateValue = 8266   '定位值
Public Const conMenu_Queue_LocateType = 8267    '定位类型
Public Const conMenu_Queue_Filter = 8268        '过滤

Public Const conMenu_Queue_PrintNumber = 8272       '打号
Public Const conMenu_Queue_InsertQueue = 8273       '插队
Public Const conMenu_Queue_RestartQueue = 8274      '重排

Public Const conMenu_Queue_Quick = 8275      '快捷窗口
Public Const conMenu_Queue_Passed = 1000      '过号



