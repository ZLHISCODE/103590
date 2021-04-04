Attribute VB_Name = "mdlPacs"
Option Explicit

''''''''插件说明''''''''''''''''''''''
'''说明：''''''''''''''''''''''''''''''''''''''''''
'''1、本例子程序中，调用PACS的部分主要在mdlPacs模块中实现。
'''2、通常调用PACS程序，涉及到三个部分：初始化，打开PACS程序（观片或者报告），结束并释放资源。
'''3、本例子以心医提供的PACS接口为例





'''''''''''''''''''''''''''''''''''''''''''''''''''''
'1、XEFORHIS.dll是心医提供给HIS医生站调用的接口文件。
''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''说明：在此处根据PACS提供的DLL，书写对应的定义语句
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function XePacsInit Lib "XEFORHIS.dll" () As Boolean
Public Declare Function XePacsCall Lib "XEFORHIS.dll" ( _
    ByVal nPatientIDType As Long, _
    ByVal lpszID As String, _
    ByVal nCallType As Long _
) As Boolean

Public Declare Function XePacsRelease Lib "XEFORHIS.dll" ()



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''说明：此处根据PACS的功能，定义对应的公共变量，方便在类模块clsPlugIn中的调用
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const gstrFunc_PACS影像调阅 = "PACS影像调阅"
Public Const gstrFunc_PACS报告调阅 = "PACS报告调阅"


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''说明：此处定义需要使用的公共变量
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private blnInitPacsConnection As Boolean        '是否需要初始化PACS连接，PACS连接只需要初始化一次



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''说明：此处编写调用PACS接口的具体过程和函数
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function InitPacs() As Boolean
'初始化跟PACS的数据库连接

'初始化心医的PACS数据库连接

    Dim blnErr As Boolean
    
    On Error GoTo err
    
    InitPacs = False
    
    blnErr = XePacsInit


    If blnErr = False Then
        MsgBox "初始化数据错误", vbOKOnly, "PACS影像接口"
        Exit Function
    End If
    
    InitPacs = True
    Exit Function
err:
    MsgBox err.Description, vbOKOnly, "PACS影像接口"
    err.Clear
End Function

Public Function ShowPacsViewer(ByVal varKeyId As Variant, lngViewType As Integer) As Boolean
'说明：显示PACS的图像浏览器和报告浏览器，根据PACS提供的参数设置本过程的参数数量和内容

'调用心医的XePacsCall函数，从IE显示WEB版本的PACS图像浏览器
'参数： varKeyId --- 医嘱ID
'       lngViewType --- 浏览方式，1为查看图像，2为查看报告
    Dim blnErr As Boolean
    
    
    On Error GoTo err
    
    ShowPacsViewer = False
    
    '先初始化
    '只针对门诊医生站和住院医生站才初始化PACS调阅图像的插件
    If blnInitPacsConnection = False Then
        Dim lngWait As Long
        blnInitPacsConnection = InitPacs
            
        '循环只是为了延时，心医的接口初始化之后直接调用图像，会提示错误，需要有一个延时
        For lngWait = 1 To 6000
        
        Next lngWait
        
    End If
        
        
    'XePacsCall 参数说明：  nPatintIDType 编号类型：1：门诊号；2：住院号；3：申请单号
    '                       nCallType 调用类型：1：查看图像；2：查看报告
    '调用过XePacsInit后，即可调用本函数来查看图像或报告
    
    If blnInitPacsConnection = True Then
        blnErr = XePacsCall(3, CStr(varKeyId), lngViewType)
        If blnErr = False Then
            MsgBox IIf(lngViewType = 1, "调用图像发生错误", "调阅报告发生错误"), vbOKOnly, "PACS影像接口"
            Exit Function
        End If
    
        ShowPacsViewer = True
    End If
    
    Exit Function
err:
    MsgBox err.Description, vbOKOnly, "PACS影像接口"
    err.Clear
End Function

Public Function PacsRelease()
'说明： 结束PACS浏览，释放连接，只需要调用一次

'调用心医XePacsRelease函数，释放数据库连接
    On Error GoTo err
    
    If blnInitPacsConnection = True Then
        XePacsRelease
    End If
    
    Exit Function
err:
   MsgBox err.Description, vbOKOnly, "PACS影像接口"
    err.Clear
End Function


