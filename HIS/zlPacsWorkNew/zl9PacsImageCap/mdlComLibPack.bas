Attribute VB_Name = "mdlComLibPack"
Option Explicit

Public gobjComLib As Object 'zl9ComLib.clsComLib


Public Sub zlCL_CloseWindow()
'关闭使用窗口
    If gobjComLib Is Nothing Then Exit Sub
    
    Call gobjComLib.CloseWindows
End Sub

Public Function zlCL_GetCacheDir() As String
'获取缓存目录
    zlCL_GetCacheDir = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
End Function


Public Function zlCL_GetResourceDir() As String
'获取资源目录
    zlCL_GetResourceDir = IIf(Len(App.Path) > 3, App.Path & "\..\附加文件\", App.Path & "..\附加文件\")
End Function

Public Function zlCL_GetPara(ByVal varPara As Variant, _
                        Optional ByVal lngSys As Long, _
                        Optional ByVal lngModual As Long, _
                        Optional ByVal strDefault As String, _
                        Optional ByVal arrControl As Variant, _
                        Optional ByVal blnSetup As Boolean, _
                        Optional intType As Integer) As String
'获取参数设置
    zlCL_GetPara = ""
    If gobjComLib Is Nothing Then Exit Function
    
    zlCL_GetPara = gobjComLib.zlDataBase.GetPara(varPara, lngSys, lngModual, strDefault, arrControl, blnSetup, intType)
End Function


Public Sub zlCL_ShowFlash(Optional strNote As String, Optional frmParent As Object)
'Flash窗口显示
    If gobjComLib Is Nothing Then Exit Sub
    
    Call gobjComLib.zlCommFun.ShowFlash(strNote, frmParent)
End Sub

Public Sub zlCL_StopFlash()
'终止Flash窗口显示
    If gobjComLib Is Nothing Then Exit Sub
    
    Call gobjComLib.zlCommFun.StopFlash
End Sub


Public Sub zlCL_ExecuteProcedure(strSQL As String, ByVal strFormCaption As String)
'执行过程
    If gobjComLib Is Nothing Then Exit Sub
    
    Call gobjComLib.zlDataBase.ExecuteProcedure(strSQL, strFormCaption)
End Sub

Public Function zlCL_GetDBObj() As Object
'获取数据库执行对象
'由于不能对zldatabase的executedreProcedure进行封装，因此只能由对象单独执行
    Set zlCL_GetDBObj = Nothing
    If gobjComLib Is Nothing Then Exit Function
    
    Set zlCL_GetDBObj = gobjComLib.zlDataBase
End Function

Public Function zlCL_GetNodeNo() As String
'获取站点编号
    If gobjComLib Is Nothing Then Exit Function
    
    zlCL_GetNodeNo = gobjComLib.gstrNodeNo
End Function

Public Function zlCL_GetPubIcons() As Object
'获取公共图标对象
    Set zlCL_GetPubIcons = Nothing
    If gobjComLib Is Nothing Then Exit Function
    
    Set zlCL_GetPubIcons = frmPubIcons.imgPublic.Icons
End Function

Public Function zlCL_Currentdate() As Date
'从数据库服务器中获取当前时间
    If gobjComLib Is Nothing Then Exit Function
    
    zlCL_Currentdate = gobjComLib.zlDataBase.Currentdate
End Function

Public Function zlCL_GetPrivFunc(lngSys As Long, lngProgId As Long) As String
'获取功能权限
    If gobjComLib Is Nothing Then Exit Function
    
    zlCL_GetPrivFunc = gobjComLib.GetPrivFunc(lngSys, lngProgId)
End Function

Public Sub zlCL_CboSetIndex(ByVal hWnd_combo As Long, ByVal lngIndex As Long)
'设置combobox的listindex
    If gobjComLib Is Nothing Then Exit Sub
    
    Call gobjComLib.zlControl.CboSetIndex(hWnd_combo, lngIndex)
End Sub


Public Sub zlCL_PressKey(bytKey As Byte)
'模拟按键发送
    If gobjComLib Is Nothing Then Exit Sub
    
    Call gobjComLib.zlCommFun.PressKey(bytKey)
End Sub


Public Function zlCL_SetPara(ByVal varPara As Variant, _
                            ByVal strValue As String, _
                            Optional ByVal lngSys As Long, _
                            Optional ByVal lngModual As Long, _
                            Optional ByVal blnSetup As Boolean = True) As Boolean
'保存参数数据
    zlCL_SetPara = False
    If gobjComLib Is Nothing Then Exit Function
    
    zlCL_SetPara = gobjComLib.zlDataBase.SetPara(varPara, strValue, lngSys, lngModual, blnSetup)
End Function


Public Function zlCL_RestoreWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'恢复窗口状态
    zlCL_RestoreWinState = False
    If gobjComLib Is Nothing Then Exit Function
    
    zlCL_RestoreWinState = gobjComLib.RestoreWinState(objForm, strProjectName, strUserDef)
End Function


Public Function zlCL_SaveWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'保存窗口状态
    zlCL_SaveWinState = False
    If gobjComLib Is Nothing Then Exit Function
    
    zlCL_SaveWinState = gobjComLib.SaveWinState(objForm, strProjectName, strUserDef)
End Function


Public Function zlCl_RegCheck(Optional blnTemp As Boolean) As String
'注册检查
    zlCl_RegCheck = ""
    If gobjComLib Is Nothing Then Exit Function
    
    zlCl_RegCheck = gobjComLib.zlRegCheck(blnTemp)
End Function

Public Function zlCL_GetRegInfo(strItem As String, Optional blnTemp As Boolean, Optional intBits As Integer) As String
'获取注册信息
    zlCL_GetRegInfo = ""
    If gobjComLib Is Nothing Then Exit Function
    
    zlCL_GetRegInfo = gobjComLib.zlRegInfo(strItem, blnTemp, intBits)
End Function


Public Sub zlCL_InitCommon(cnMain As ADODB.Connection)
'初始化连接对象
    If gobjComLib Is Nothing Then Exit Sub
    
    Call gobjComLib.InitCommon(cnMain)
    
    Call gobjComLib.SetDbUser(cnMain.Properties.Item("User ID").value)
End Sub





