Attribute VB_Name = "mdlCardEx"
Option Explicit

Public gcnOracle As ADODB.Connection '公共连接对象

Public Enum Enum_Modue '模块号
    m门诊医嘱模块 = 1252
    m住院医嘱模块 = 1253
    m住院护士站模块 = 1254
    m临床路径模块 = 1256
    m病历模块 = 1070
    m人员管理模块 = 1002
    m医嘱附费模块 = 1257
    
    m门诊医生工作站 = 1260
    m住院医生工作站 = 1261
    m住院护士工作站 = 1262
    m医技工作站 = 1263
    m新版护士站 = 1256
End Enum


