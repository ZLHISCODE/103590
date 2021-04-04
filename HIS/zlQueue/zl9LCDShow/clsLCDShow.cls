VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLCDShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim mfrmLCDShow As frmLCDShow
Dim mfrmDrugShow As frmDrugShow

Public Function zlShow(cnOracle As ADODB.Connection, str队列名称() As String, _
    Optional str诊室 As String = "", Optional str医生 As String = "", _
    Optional strExcludeData As String = "", Optional intViewDataType As Integer = 0, _
    Optional blnComeBackFirst As Boolean = True) As String
    
    
    '在电视屏幕中显示排队情况
    '返回值为排队情况字符串
    
    On Error GoTo err
    
    '判断是否需要创建显示窗口
    If mfrmLCDShow Is Nothing Then
        '创建显示窗口
        Set mfrmLCDShow = New frmLCDShow
        
        Call mfrmLCDShow.SetFaceFont
        Call mfrmLCDShow.SetFacePostion
        mfrmLCDShow.zlShowMe cnOracle, str队列名称, str诊室, str医生, strExcludeData, intViewDataType, blnComeBackFirst
    Else
        Call mfrmLCDShow.zlSetPara(str队列名称, str诊室, str医生, strExcludeData, blnComeBackFirst)
    End If
    
    
    
    '构建返回的排队情况字符串
    zlShow = "" 'mfrmLCDShow.lblShowText.Caption
    
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlDrugShow(ByVal lngStoreID As Long, ByVal strWin As String, ByVal bln配药 As Boolean, ByVal bln配药确认 As Boolean, Optional ByVal strName As String) As String
    '药房排队叫号显示输出
    '返回值为当前呼叫内容
    Dim intMode As Integer '窗口模式：0-单窗口；1-多窗口
    Dim strTemp As String
    Dim strReg As String
    
    On Error GoTo err
    
    strReg = "公共模块\药房排队叫号\液晶电视"
    
    intMode = Val(GetSetting("ZLSOFT", "公共模块\药房排队叫号\液晶电视", "窗口模式", "0"))
    
    If intMode < 0 Or intMode > 1 Then
        intMode = 0
    End If
    
    If intMode = 0 Then
        strTemp = strWin
    End If
    
    If strName = "" Then Set mfrmDrugShow = Nothing
    If mfrmDrugShow Is Nothing Then
        '创建显示窗口
        Set mfrmDrugShow = New frmDrugShow
        
        '更新窗口设置
        
        '显示窗口
        mfrmDrugShow.ShowMe lngStoreID, strTemp, bln配药, bln配药确认
    Else
        '窗口已创建时，更新当前呼叫栏
        mfrmDrugShow.ChangeCall strWin, strName
    End If
        
    
    '返回当前的呼叫内容
    zlDrugShow = ""
    
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zlClose() As Boolean
    '关闭液晶电视的显示窗口
    
    On Error GoTo err
    
    If mfrmLCDShow Is Nothing Then Exit Function
    
    Unload mfrmLCDShow
    Set mfrmLCDShow = Nothing
    
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlDrugShowClose() As Boolean
    '关闭液晶电视的显示窗口
    
    On Error GoTo err
    
    If mfrmDrugShow Is Nothing Then Exit Function
    
    Unload mfrmDrugShow
    Set mfrmDrugShow = Nothing
    
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function zlSetup(frmParent As Object) As Boolean
    '设置液晶屏的相关参数
    
    '参数设置，使用模式窗体来显示
    zlSetup = frmLCDSetup.zlShowMe(frmParent)
    
    If mfrmLCDShow Is Nothing Then Exit Function
    
    Call mfrmLCDShow.SetFaceFont
    Call mfrmLCDShow.SetFacePostion
    
End Function

Public Function zlDrugSetup(frmParent As Object, ByVal strSendWindows As String) As Boolean
    '设置药房排队叫号显示输出设备的相关参数
    
    '参数设置，使用模式窗体来显示
    zlDrugSetup = frmDrugSetup.ShowMe(strSendWindows, frmParent)
    
    If mfrmDrugShow Is Nothing Then Exit Function
    
    '更新窗口设置
    Call mfrmDrugShow.SetFacePostion
    Call mfrmDrugShow.SetFont
    
End Function
