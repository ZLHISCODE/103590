VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public lngTestCount As Long



Public Sub InitActiveVideo()
'使其触发Init事件...

End Sub

''获取虚拟键名称
'Public Function GetKeyAliasEx(ByVal lngVirtualKey As Long) As String
'    GetKeyAliasEx = mdlPublic.GetKeyAliasEx(lngVirtualKey)
'End Function

''获取组合键名称
'Public Function GetKeyAlias(ByVal lngKeyCode As Integer, ByVal lngShift As Integer) As String
'    GetKeyAlias = mdlPublic.GetKeyAlias(lngKeyCode, lngShift)
'End Function


Public Function GetInstanceCount() As Boolean
'获取已经创建的实例数量
    GetInstanceCount = glngInstanceCount
End Function


Public Sub ForceCreate()
'强制重新创建资源
    Call CreateGlobalResource
End Sub


Public Sub ForceFree()
'强制释放资源
    Call FreeGlobalResource
End Sub


Private Sub CreateGlobalResource()
'初始化视频采集对象
    
BUGEX "Global>CreateResource 1", True

    gblnOpenDebug = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "IsOpenDebug", 0)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "IsOpenDebug", IIf(gblnOpenDebug, 1, 0))
    
    If gblnOpenDebug Then
BUGEX "Global>CreateResource Open Debug Successed", True
    Else
BUGEX "Global>CreateResource Open Debug Failed", True
    End If
        
    '创建公共资源
    If gobjCapturePar Is Nothing Then
        BUGEX "Global>CreateResource gobjCapturePar", True
        Set gobjCapturePar = New clsCaptureParameter
    End If
    
'BUGEX "CreateResource 2", True
'    If gobjComLib Is Nothing Then
'        BUGEX "CreateResource gobjComLib", True
'
'        Set gobjComLib = CreateObject("zl9ComLib.clsComLib")  'New zl9ComLib.clsComLib  '
'    End If
    

BUGEX "Global>CreateResource 2", True
    If gobjVideo Is Nothing Then
        BUGEX "Global>CreateResource gobjVideo", True
        Set gobjVideo = New frmWork_Video
    End If
    
BUGEX "Global>CreateResource 3", True
    If gobjNotifyEvent Is Nothing Then
        BUGEX "Global>CreateResource gobjNotifyEvent", True
        Set gobjNotifyEvent = New clsNotifyEvent
    End If

BUGEX "Global>CreateResource End.", True

End Sub


Private Sub FreeGlobalResource()
'释放视频采集资源
        
BUGEX "Global>FreeGlobalResource 1", True
        '释放事件通知对象
        If Not gobjNotifyEvent Is Nothing Then
            Set gobjNotifyEvent = Nothing
        End If
                
BUGEX "Global>FreeGlobalResource 2", True
        '只有实例数量为一的时候，才释放全局公共资源
        If Not gobjVideo Is Nothing Then
            Unload gobjVideo
            Set gobjVideo = Nothing
        End If
        
BUGEX "Global>FreeGlobalResource 3", True
        If Not gobjZOrder Is Nothing Then
            Call gobjZOrder.RemoveAll
            Set gobjZOrder = Nothing
        End If
        
BUGEX "Global>FreeGlobalResource 4", True

        '设置参数配置对象
        If Not gobjCapturePar Is Nothing Then
            Set gobjCapturePar = Nothing
        End If
        
BUGEX "Global>FreeGlobalResource 5", True
        '释放oracle对象
        If Not gcnVideoOracle Is Nothing Then
            Set gcnVideoOracle = Nothing
        End If
        
        '释放图像资源
        Unload frmPubIcons
        
BUGEX "Global>FreeGlobalResource 6", True
        If Not gobjComLib Is Nothing Then
            Call zlCL_CloseWindow
            Set gobjComLib = Nothing
        End If
    
BUGEX "Global>FreeGlobalResource End.", True
End Sub



Private Sub Class_Initialize()
'打开调试配置
BUGEX "Global Initialize 1", True
    Call CreateGlobalResource
    
BUGEX "Global Initialize End", True
End Sub

Private Sub Class_Terminate()
'验证全局对象是否被卸载
BUGEX "Global Terminate 1", True

    Call FreeGlobalResource
    
BUGEX "Global Terminate End", True
End Sub
