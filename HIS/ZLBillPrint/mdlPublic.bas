Attribute VB_Name = "mdlPublic"
Option Explicit

Public gcnOracle As ADODB.Connection        'HIS调用本部件时传入的公共数据库连接
Public gstrUserCode As String   '当前操作员编号
Public gstrUserName As String   '当前操作员姓名

Public Enum ModulNO
    FOutBillPrint = 1121    '门诊收费
    FInBillPrint = 1137     '住院结帐
End Enum
Public glngSys As Long          '当前调用系统编号，100=ZLHIS标准版
Public glngModul As ModulNO        '当前调用模块号，1121=门诊收费,1137=住院结帐

