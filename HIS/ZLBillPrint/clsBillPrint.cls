VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBillPrint"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'说明：
'    必须引用数据访问部件ms ado 6.0:msado15.dll，可通过ZLHIS基础安装包安装，或单独安装MDAC_TYP_V2.6
'    类模块作为通用的外部接口调用，每个地区的票据打印单独封装在一个独立的模块文件中
'    mdlBillPrint_BJ=北京票据打印接口
'注意：
'   如果第三方接口是C或C#等语言编写的，参数类型为字节数组型时，VB中需要定义为定长字符串型，并且先进行初始化
'   例如：dim strError As String * 255      strError = String(255, " ")

Public Function zlInitialize(ByRef cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal lngModul As Long, _
    ByVal strUserCode As String, ByVal strUserName As String) As Boolean
'功能：初始化接口,在HIS进入模块时调用(例如：进入收费管理界面)，可在本接口中完成第三方票据打印接口的初始化、登录、打开连接等调用

'参数：cnMain=数据库连接对象
'      strUserCode=当前操作员编号
'      strUserName=当前操作员姓名
'      lngSys=当前调用系统编号，100=ZLHIS标准版
'      lngModul=当前调用模块号，1121=门诊收费,1137=住院结帐
'
'返回：初始化成功/失败
        
    Set gcnOracle = cnMain
    glngSys = lngSys
    glngModul = lngModul
    gstrUserCode = strUserCode
    gstrUserName = strUserName
    
    zlInitialize = Init
End Function

Public Function zlTerminate() As Boolean
'功能：终止接口，在HIS退出模块时调用(例如：退出收费管理界面)，可在本接口中完成第三方票据打印接口的资源释放、断开连接等调用
'返回：执行成功/失败
    
    zlTerminate = Term
End Function

Public Function zlConfigure() As Boolean
'功能：参数设置,在HIS收费或结帐模块参数设置"(文件/参数设置)中执行"票据打印设置"时调用，可在本接口中完成第三方票据打印接口的参数设置、配置更改等调用。
'返回：执行成功/失败
    
    zlConfigure = SYSConfigure
End Function

Public Function zlDiscardBill(ByVal lng领用ID As Long, ByVal lng票种 As Long, ByVal str票号前缀 As String, _
    ByVal str开始票号 As String, ByVal str结束票号 As String, ByVal DateAdd As Date, ByVal str报损人 As String) As Boolean
'功能：票据报损，保存完成后调用本接口
'参数：lng领用ID = 票据领用记录.ID，即领用批次ID
'      lng票种   = 1-收费票据,2-预交票据,3-结帐票据,4-挂号票据,5-就诊卡
'      DateAdd   = 报损时间
'      str报损人 = 报损人姓名

'返回：执行成功/失败
    
    zlDiscardBill = DiscardBill(lng领用ID, lng票种, str票号前缀, str开始票号, str结束票号, DateAdd, str报损人)
End Function

Public Function zlPrintBill(ByVal strNOs As String, ByVal lngBalanceId As Long) As Boolean
'功能：门诊收费或住院结帐完成后调用票据打印
'      门诊修改一张单据号，先调用票据作废接口zlEraseBill，再调用本接口。
'参数：strNOs=门诊收费：以逗号分隔的带引号的多个单据号(一次打印单张或多张单据):'F0000001','F0000002',...
'      lngBalanceId=住院结帐：结帐单ID
'返回：执行成功/失败
    If glngModul = FOutBillPrint Then
        '门诊收费
        zlPrintBill = PrintBillOut(strNOs)
    Else
        '住院结帐
        zlPrintBill = PrintBillIn(lngBalanceId)
    End If
End Function


Public Function zlRePrintBill(ByVal strNOs As String, ByVal lngBalanceId As Long, ByVal strInvoice As String) As Boolean
'功能：重新打印票据，包括直接使用重打功能，以及门诊部分退费时重打两种情况。
'       门诊部分退费重打时，先调用zlEraseBill接口作废票据，再调用本接口。
'       门诊全部退费时仅收回票据，不调用本接口，只调用作废接口。

'参数：strNOs=门诊收费：以逗号分隔的带引号的多个单据号(一次打印单张或多张单据):'F0000001','F0000002',...
'      lngBalanceId=住院结帐：结帐单ID
'      strInvoice=本次重打使用的起始票据号
'返回：执行成功/失败
    If glngModul = FOutBillPrint Then
        '门诊收费
        zlRePrintBill = RePrintBillOut(strNOs, strInvoice)
    Else
        '住院结帐
        zlRePrintBill = RePrintBillIn(lngBalanceId, strInvoice)
    End If
End Function


Public Function zlEraseBill(ByVal strNOs As String, ByVal lngBalanceId As Long) As Boolean
'功能：作废已打印票据，门诊全退和住院结帐作废时调用
'      门诊修改单据重打前和部分退费重打前先调用本接口。
'参数：strNOs=门诊收费：以逗号分隔的带引号的多个单据号(一次打印单张或多张单据):'F0000001','F0000002',...
'      lngBalanceId=住院结帐：结帐单ID
'返回：执行成功/失败
    If glngModul = FOutBillPrint Then
        '门诊收费
        zlEraseBill = EraseBillOut(strNOs)
    Else
        '住院结帐
        zlEraseBill = EraseBillIn(lngBalanceId)
    End If
End Function

