Attribute VB_Name = "mdlCommunity"
Option Explicit

Public gcnOracle As ADODB.Connection '全局数据库链接
Public gstrSysName As String '用于消息提示框

Public grsCommunity As ADODB.Recordset '社区目录缓存
Public gcolCommunity As New Collection '社区部件集合
Public gobjCommunity As Object '当前使用的社区部件

Public Function GetCommunity(ByVal int社区 As Integer) As Object
'功能：动态初始化指定的社区部件和环境，并返回对应的社区部件
'返回：如果初始化成功则返回社区部件对象
    Dim objTemp As Object
    
    '取集合中已初始化好的社区部件
    On Error Resume Next
    Set objTemp = gcolCommunity("_" & int社区)
    Err.Clear: On Error GoTo 0
    
    '如果没有表示还没有初始化
    If objTemp Is Nothing Then
        grsCommunity.Filter = "序号=" & int社区
        If grsCommunity.EOF Then Exit Function '因为社区有外键，应该不会出现这种情况
        If zlCommFun.Nvl(grsCommunity!启用, 0) = 0 Then
            MsgBox grsCommunity!名称 & "当前没有启用。", vbExclamation, gstrSysName
            Exit Function
        End If
        
        '创建该社区部件
        On Error Resume Next
        Set objTemp = CreateObject(grsCommunity!部件名 & ".clsCommunity")
        If Err.Number <> 0 Then
            MsgBox grsCommunity!名称 & "部件""" & grsCommunity!部件名 & ".dll""没有正确安装。", vbExclamation, gstrSysName
            Err.Clear: Exit Function
        End If
        
        '初始化该社区部件
        Err.Clear: On Error GoTo errH
        If Not objTemp.Initialize(gcnOracle) Then Exit Function
        
        '初始化成功之后加入部件集合
        gcolCommunity.Add objTemp, "_" & int社区
    End If
    
    Set GetCommunity = objTemp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
