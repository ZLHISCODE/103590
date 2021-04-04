Attribute VB_Name = "mdlMUSE"
Option Explicit

''''''''插件说明''''''''''''''''''''''
'''说明：''''''''''''''''''''''''''''''''''''''''''
'''1、本例子程序中，调用MUSE心电系统检查结果的部分主要在mdlMUSE模块中实现。
'''2、通常调用MUSE程序，直接调用本机浏览器，打开心电检查结果的链接。


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''说明：此处根据MUSE的功能，定义对应的公共变量，方便在类模块clsPlugIn中的调用
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const gstrFunc_MUSE心电结果调阅 = "心电结果调阅"

Public Function ShowMUSEViewer(ByVal varKeyId As Variant) As Boolean
'说明：显示MUSE的浏览器

'参数： varKeyId --- 医嘱ID

    Dim blnErr As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strURL As String
    Dim i As Integer
    
    On Error GoTo err
    
    ShowMUSEViewer = False
    
    '从HIS的数据库中查找本次医嘱ID对应的心电系统连接URL
    strSQL = "Select 执行说明 From 病人医嘱发送 Where 医嘱ID = " & varKeyId
    Set rsTemp = gcnOracle.Execute(strSQL)
    
    '因为只知道医嘱ID，不知道发送号，对于长嘱，需要循环查找第一个有执行说明的记录，用来打开心电系统的检查结果
    For i = 1 To rsTemp.RecordCount
        strURL = IIf(IsNull(rsTemp!执行说明), "", rsTemp!执行说明)
        If strURL <> "" Then
            Exit For
        End If
        rsTemp.MoveNext
    Next i
    
    If strURL <> "" Then
        '打开浏览器
        Shell "explorer " & strURL, 0
        ShowMUSEViewer = True
    End If
    
    Exit Function
err:
    MsgBox err.Description, vbOKOnly, "MUSE接口错误"
    err.Clear
End Function
