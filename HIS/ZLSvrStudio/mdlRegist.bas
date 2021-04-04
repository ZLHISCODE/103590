Attribute VB_Name = "mdlRegist"
Option Explicit


Public Function zlHomePage(hwnd As Long) As Boolean
'功能：根据产品发行码，联结主页
    Dim strCode As String
    
    strCode = gobjRegister.zlRegInfo("支持商URL")
    If strCode <> "-" Then
        ShellExecute hwnd, "open", "http://" & strCode, "", "", 1
        zlHomePage = True
    End If
End Function

Public Function zlWebForum(hwnd As Long) As Boolean
'功能：根据产品发行码，连接论坛
    Dim strCode As String
    
    strCode = "www.zlsoft.com/techbbs/index.asp"
    If strCode <> "-" Then
        ShellExecute hwnd, "open", "http://" & strCode, "", "", 1
        zlWebForum = True
    End If
End Function

Public Sub ShowAbout(Optional frmParent As Object)
    Dim frmShow As New frmAbout
    If frmParent Is Nothing Then
        frmShow.Show 1
    Else
        Load frmShow
        err.Clear
        On Error Resume Next
        frmShow.Show 1, frmParent
        If err.Number <> 0 Then
            err.Clear
            frmShow.Show 1
        End If
    End If
End Sub

Public Function zlMailTo(hwnd As Long) As Boolean
'功能：根据产品发行码发送电子邮件
    Dim strCode As String
    strCode = gobjRegister.zlRegInfo("支持商MAIL")
    If strCode <> "-" Then
        ShellExecute hwnd, "open", "mailto:" & strCode, "", "", 1
        zlMailTo = True
    End If
End Function

Public Function zlGetRegSystems() As ADODB.Recordset
'功能：获取已经注册的系统
    Dim strSQL As String, rsSys As New ADODB.Recordset
    On Error GoTo errH
    strSQL = "Select * From zlSystems S Where Trunc(S.编号 / 100) In (Select Distinct R.系统 From zlRegFunc R Where R.功能 = '基本') order by s.编号 "
    Call OpenRecordset(rsSys, strSQL, "zlGetRegSystems")
    Set zlGetRegSystems = rsSys
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "获取已经注册系统失败，失败信息：" & err.Description, vbInformation, App.Title
End Function

Public Function RunRegistFile(ByVal objParent As Object, ByVal cnnTools As ADODB.Connection, ByVal strPassword As String, ByVal strServer As String, ByVal strRegFunFile As String) As Boolean
'功能：以SQLPlus执行注册码函数创建文件
    Dim objScript As clsRunScript
    
    Set objScript = New clsRunScript
    With objScript
        Set .Connection = cnnTools: .ConnectType = 1
        Call .InitGlobalPara(objParent)
        Call .InitUserList(, , strPassword)
        .Server = strServer
        If .OpenFile(strRegFunFile) = False Then
            Exit Function
        End If
        
        Do While Not .EOF
            If .SQLInfo.PartSQL <> "EXIT" Then
                If Not .ExecuteSQL(.SQLInfo) Then
                    Exit Function
                End If
            End If
            .ReadNextSQL
        Loop
    End With
    RunRegistFile = True
End Function
