Attribute VB_Name = "mdlScriptTool"
Option Explicit


'系统参数信息
Public Type SYSPARAM_INFO
    费用金额小数位数 As String
    收费诊疗项目匹配 As String
    结帐票据号长度 As Integer
    收费票据号长度 As Integer
    就诊卡号码长度 As Integer
    就诊卡字母前缀 As String
    就诊卡密文显示 As Boolean
    项目输入匹配方式 As Integer '0-双向;1-从左
    系统号 As Long
    系统名称 As String
    产品名称 As String
    模块号 As Long
    所有者 As String
    收费票种 As Integer
    结帐票种 As Integer
    结帐票号严格控制 As Boolean
    收费票号严格控制 As Boolean
    连接HIS报告 As Byte
End Type

'----------------------------------------------------------------------------------------------------------------------
'全局变量申明

Public ParamInfo As SYSPARAM_INFO
Public glngTXTProc As Long                              '保存默认的消息函数的地址

Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
    '******************************************************************************************************************
    '功能：设置指定的参数值
    '参数：varPara=参数号或参数名，以数字或字符类型传入区分
    '      strValue=要设置的参数值
    '      lngModual=使用该参数的模块号，如1230
    '      blnPrivate=该参数是否用户私有参数
    '返回：设置是否成功
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    GetPara = zlDatabase.GetPara(varPara, ParamInfo.系统号, lngModual, strDefault, blnNotCache)

errHand:

End Function

Public Function SetPara(ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：设置指定的参数值
    '参数：varPara=参数号或参数名，以数字或字符类型传入区分
    '      strValue=要设置的参数值
    '      lngModual=使用该参数的模块号，如1230
    '      blnPrivate=该参数是否用户私有参数
    '返回：设置是否成功
    '******************************************************************************************************************

    On Error GoTo errH
        
    SetPara = zlDatabase.SetPara(varPara, strValue, ParamInfo.系统号, lngModual, blnSetup)

    Exit Function
    
errH:

End Function


Public Function CommandBarUpdatePublic(Control As Object, frmMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Select Case Control.Id
    Case conMenu_View_ToolBar_Button            '工具栏
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = frmMain.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              '图标文字
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = Not (frmMain.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size              '大图标
        Control.Checked = frmMain.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar                 '状态栏
        Control.Checked = frmMain.stbThis.Visible
    End Select
    
    CommandBarUpdatePublic = True
End Function
