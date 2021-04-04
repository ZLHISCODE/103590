VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICCardDev_JCSRFID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsCardDev

'读卡函数声明
Private Declare Function piccreadex Lib "OUR_MIFARE.dll" (ByVal ctrlword As Byte, ByVal serial As Long, ByVal area As Byte, ByVal keyA1B0 As Byte, ByVal picckey As Long, ByVal piccdata0_2 As Long) As Byte

'写卡函数声明
Private Declare Function piccwriteex Lib "OUR_MIFARE.dll" (ByVal ctrlword As Byte, ByVal serial As Long, ByVal area As Byte, ByVal keyA1B0 As Byte, ByVal picckey As Long, ByVal piccdata0_2 As Long) As Byte

'修改单区函数声明
Private Declare Function piccchangesinglekey Lib "OUR_MIFARE.dll" (ByVal ctrlword As Byte, ByVal serial As Long, ByVal area As Byte, ByVal keyA1B0 As Byte, ByVal piccoldkey As Long, ByVal piccnewkey As Long) As Byte

'让设备发出声响函数声明
Private Declare Function pcdbeep Lib "OUR_MIFARE.dll" (ByVal xms As Long) As Byte

'读取设备编号函数声明
Private Declare Function pcdgetdevicenumber Lib "OUR_MIFARE.dll" (ByVal devicenumber As Long) As Byte

'控制字定义,控制字指定,控制字的含义请查看本公司网站提供的动态库说明
Private Const BLOCK0_EN = &H1
Private Const BLOCK1_EN = &H2
Private Const BLOCK2_EN = &H4
Private Const NEEDSERIAL = &H8
'Private Const EXTERNKEY = &H10
Private Const NEEDHALT = &H20
        

Private Function clsCardDev_CloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    clsCardDev_CloseCard = True
End Function

Private Function clsCardDev_GetSpare() As Currency
    
End Function

Private Function clsCardDev_initCard(Optional strInput As String, Optional strOut As String) As Boolean
    clsCardDev_initCard = True
End Function

Private Function clsCardDev_PaymentSwap(ByVal dblMoney As Double, dblBalance As Double, intCardType As Integer, ByVal lngHOSCode As Long, strCardNO As String, strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean
    clsCardDev_PaymentSwap = True
End Function

Private Function clsCardDev_ReadCard(Optional strInput As String, Optional blnCancel As Boolean = False) As String
    Dim status As Byte '存放返回值
    
    Dim myareano As Byte '区号
    Dim authmode As Byte '密码类型，用A密码或B密码
    Dim myctrlword As Byte '控制字
    Dim mypicckey(0 To 5) As Byte '密码
    Dim mypiccserial(0 To 3) As Byte '卡序列号
    Dim mypiccdata(0 To 47) As Byte '卡数据缓冲
    
    Dim strMyPicdata As String  '接收返回字符串
    Dim i As Integer   '循环变量
    Dim strBit As String
    Dim strChar As String
    
    On Err GoTo errH:
    
    '控制字指定,控制字的含义请查看本公司网站提供的动态库说明
    myctrlword = BLOCK0_EN + BLOCK1_EN + BLOCK2_EN
    
    '指定区号
    myareano = 8 '指定为第8区
    '批定密码模式
    authmode = 1 '大于0表示用A密码认证，推荐用A密码认证
    
    '指定密码
    mypicckey(0) = &HFF
    mypicckey(1) = &HFF
    mypicckey(2) = &HFF
    mypicckey(3) = &HFF
    mypicckey(4) = &HFF
    mypicckey(5) = &HFF
    
    status = piccreadex(myctrlword, VarPtr(mypiccserial(0)), myareano, authmode, VarPtr(mypicckey(0)), VarPtr(mypiccdata(0)))
    '在下面设定断点，然后查看mypiccserial、mypiccdata，
    '调用完 piccreadex函数可读出卡序列号到 mypiccserial，读出卡数据到mypiccdata，
    '开发人员根据自己的需要处理mypiccserial、mypiccdata 中的数据了。
    '处理返回函数
    If gbln自动读取 Then
        For i = 17 To 20
         strBit = Hex(mypiccdata(i))
         If Len(strBit) = 1 Then strBit = "0" & strBit
            strChar = strBit
            strMyPicdata = strMyPicdata & UCase(strChar)
        Next
        If strMyPicdata = "00000000" Then
           strMyPicdata = ""
        End If
    Else
        Select Case status
            Case 0:
                For i = 17 To 20
                    strBit = Hex(mypiccdata(i))
                    If Len(strBit) = 1 Then strBit = "0" & strBit
                    strChar = strBit
                    strMyPicdata = strMyPicdata & UCase(strChar)
                Next
                If strMyPicdata = "00000000" Then
                    strMyPicdata = ""
                End If
            Case 1:
                MsgBox "刷卡太快，导致块数据读取失败！", vbInformation
            Case 2, 3:
                MsgBox "刷卡太快，导致块数据读取不完整！", vbInformation
            Case 8:
                MsgBox "请将卡放在感应区！", vbInformation
            Case 9:
                MsgBox "有多张卡在感应区，寻卡过程中防冲突失败！", vbInformation
            Case 10:
                MsgBox "对不起，该卡已休眠，无法选中！", vbInformation
            Case 11:
                MsgBox "密码装载失败！", vbInformation
            Case 12:
                MsgBox "密码验证失败！", vbInformation
            Case 21 '没有动态库
                MsgBox "找不到动态库ICUSB.DLL请将ICUSB.DLL拷贝到HIS软件安装后的目录APPSOFT下", vbInformation
            Case 22:
                MsgBox "动态库或驱动程序出现异常！", vbInformation
            Case 24:
                MsgBox "操作超时,可能是病毒导致！", vbInformation
            Case Else
                MsgBox "未知错误", vbCritical
        End Select
    End If
    If strMyPicdata <> "" Then
      Call pcdbeep(100)
    End If
    clsCardDev_ReadCard = strMyPicdata
    Exit Function
    '0   操作成功，读出的数据有效。
    '1   0~2块都没读出来，可能刷卡太块。但卡序列号已被读出来。
    '2   第0块已被读出，但1~2块读取失败。卡序列号已被读出来。
    '3   第0、1块已被读出，但2块读取失败。卡序列号已被读出来。
    '8   寻卡错误，根本就没有卡在感应区，*serial无效。
    '9   有多张卡在感应区，寻卡过程中防冲突失败，*serial无效。
    '10  该卡可能已被休眠，无法选中，但卡序列号已被读出，*serial数组中的数据有效。
    '11  密码装载失败?
    '12  密码认证失败?
    '21  本函数需要引用的动态库ICUSB.DLL不在当前目录下?
    '22  动态库或驱动程序异常，解决方法是退出程序，拔出IC卡读写器，重装驱动程序再插上IC卡读写器重试，或者重新拷贝动态库OUR_MIFARE.dll到正确的位置。
    '24  操作超时。可能是电脑中毒导致USB帧传递调度缓慢，或者是IC卡读写器有问题，解决方法是重启电脑或重新拔插IC卡读写器。
    '27  USB传输不稳定导致传输的字符不全。不需理会这个错误，因为基本上是不会出现这个错误的。
    '28  USB传输不稳定导致CRC校验错。不需理会这个错误，因为基本上是不会出现这个错误的。
    '其他 未知错误
errH:
    MsgBox Err.Description, vbExclamation
End Function

Private Function clsCardDev_ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean
    clsCardDev_ReturnSwap = True
End Function

Private Sub clsCardDev_SetCard(Optional strInput As String, Optional strOut As String)
    Call frmParameterSet.ShowMe(6)
End Sub


Private Function clsCardDev_UploadSwap(ByVal lng病人ID As Long, strOut As String) As Boolean

End Function


