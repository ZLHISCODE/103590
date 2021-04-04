Attribute VB_Name = "mdlMenu"
Option Explicit

'*********************************************************************************************************************
'
'菜单相关处理过程
'
'*********************************************************************************************************************


'查询快捷键配置
Public Sub BindMenuShortcut(ByVal strProjectName As String, ByVal lngModule As Long, objMenu As Object)
    Dim strSQL As String
    Dim rsShoftcutCfg As ADODB.Recordset
    Dim objMain As Object

    strSQL = "select id,控制键,字符键,组合名,菜单ID from 快捷功能信息 where 项目=upper([1]) and 模块号=[2]"

    Set rsShoftcutCfg = zlDatabase.OpenSQLRecord(strSQL, "绑定菜单快捷键", strProjectName, lngModule)
    
    Set objMain = objMenu
    
    Call RecursionBindMenu(objMain, objMenu.ActiveMenuBar, rsShoftcutCfg)
End Sub


'绑定菜单快捷方式(递归调用绑定快捷菜单)
Private Sub RecursionBindMenu(cbrMain As Object, objMenu As Object, rsShoftcutCfg As ADODB.Recordset)
    Dim i As Long
    
    If objMenu Is Nothing Then Exit Sub
    If objMenu.Controls.Count <= 0 Then Exit Sub
    
    For i = 1 To objMenu.Controls.Count
        Call BindMenuItemShortcut(cbrMain, objMenu.Controls.Item(i), rsShoftcutCfg)

        If objMenu.Controls.Item(i).Type = xtpControlPopup Or objMenu.Controls.Item(i).Type = xtpControlButtonPopup Then
            If objMenu.Controls.Item(i).CommandBar.Controls.Count > 0 Then
                Call RecursionBindMenu(cbrMain, objMenu.Controls.Item(i).CommandBar, rsShoftcutCfg)
            End If
        End If
    Next i
End Sub

'绑定单个菜单的快捷方式
Private Sub BindMenuItemShortcut(cbrMain As Object, cbrControl As Object, rsShoftcutCfg As ADODB.Recordset)
    If rsShoftcutCfg Is Nothing Then Exit Sub
    
    Dim lngFuncKey As Long
    Dim lngCharKey As Long
    Dim lngCommandKey As Long
    
    Dim strKeyAlias As String

    rsShoftcutCfg.Filter = "菜单ID=" & cbrControl.ID
    
    If rsShoftcutCfg.RecordCount > 0 Then
        lngFuncKey = Val(Nvl(rsShoftcutCfg!控制键))
        lngCharKey = Val(Nvl(rsShoftcutCfg!字符键))
        strKeyAlias = Nvl(rsShoftcutCfg!组合名)

        'F8固定为快捷键采集使用
        If lngFuncKey = vbKeyF8 Or lngCharKey = vbKeyF8 Then Exit Sub
        
        If (lngFuncKey <> 0 Or lngCharKey <> 0) And InStr(strKeyAlias, "MENU") <= 0 Then
            lngCommandKey = 0
 
            If (lngFuncKey And vbCtrlMask) <> 0 Then
                lngCommandKey = lngCommandKey + FCONTROL
            End If
    
            If (lngFuncKey And vbShiftMask) <> 0 Then
                lngCommandKey = lngCommandKey + FSHIFT
            End If
    
            If (lngFuncKey And vbAltMask) <> 0 Then
                lngCommandKey = lngCommandKey + FALT
            End If
            
            '绑定菜单快捷键
            Call cbrMain.KeyBindings.Add(lngCommandKey, lngCharKey, cbrControl.ID)
            
        ElseIf InStr(strKeyAlias, "MENU") > 0 Then
            If InStr(cbrControl.Caption, "(&") <= 0 Then
                cbrControl.Caption = cbrControl.Caption & "(&" & Replace(strKeyAlias, "MENU+", "") & ")"
            End If
        End If
    End If
    
End Sub



Public Sub CreateViewAndHelpMenu(objViewMenu As Object, objHelpMenu As Object, _
    Optional ByVal strMenuTag As String = "")
    
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    
    
    'Begin----------------------查看菜单--------------------------------------
    If Not (objViewMenu Is Nothing) Then
        Set cbrMenuBar = objViewMenu
        
        With cbrMenuBar.CommandBar
            Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(T)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 1
            
                With cbrControl.CommandBar '二级菜单
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(0)")
                        cbrPopControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                        
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(1)")
                        cbrPopControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                End With
            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(S)")
                cbrControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
        End With
    End If

    'Begin----------------------帮助菜单--------------------------------------默认可见
    If Not (objHelpMenu Is Nothing) Then
        Set cbrMenuBar = objHelpMenu
        
        With cbrMenuBar.CommandBar
            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_Help, "帮助主题(M)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 901
                
            Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_Help_Web, "WEB上的中联(W)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 1
                
                With cbrControl.CommandBar
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Forum, "中联论坛(0)")
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                        
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Home, "中联主页(1)")
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                        
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(2)")
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 9022
                End With
                
            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_About, "关于…(A)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 1
        End With
    End If
End Sub

'*********************************************************************************************************************
