Attribute VB_Name = "mdlRegist"
Option Explicit


Public Function zlHomePage(hwnd As Long) As Boolean
'���ܣ����ݲ�Ʒ�����룬������ҳ
    Dim strCode As String
    
    strCode = gobjRegister.zlRegInfo("֧����URL")
    If strCode <> "-" Then
        ShellExecute hwnd, "open", "http://" & strCode, "", "", 1
        zlHomePage = True
    End If
End Function

Public Function zlWebForum(hwnd As Long) As Boolean
'���ܣ����ݲ�Ʒ�����룬������̳
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
'���ܣ����ݲ�Ʒ�����뷢�͵����ʼ�
    Dim strCode As String
    strCode = gobjRegister.zlRegInfo("֧����MAIL")
    If strCode <> "-" Then
        ShellExecute hwnd, "open", "mailto:" & strCode, "", "", 1
        zlMailTo = True
    End If
End Function

Public Function zlGetRegSystems() As ADODB.Recordset
'���ܣ���ȡ�Ѿ�ע���ϵͳ
    Dim strSQL As String, rsSys As New ADODB.Recordset
    On Error GoTo errH
    strSQL = "Select * From zlSystems S Where Trunc(S.��� / 100) In (Select Distinct R.ϵͳ From zlRegFunc R Where R.���� = '����') order by s.��� "
    Call OpenRecordset(rsSys, strSQL, "zlGetRegSystems")
    Set zlGetRegSystems = rsSys
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "��ȡ�Ѿ�ע��ϵͳʧ�ܣ�ʧ����Ϣ��" & err.Description, vbInformation, App.Title
End Function

Public Function RunRegistFile(ByVal objParent As Object, ByVal cnnTools As ADODB.Connection, ByVal strPassword As String, ByVal strServer As String, ByVal strRegFunFile As String) As Boolean
'���ܣ���SQLPlusִ��ע���뺯�������ļ�
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
