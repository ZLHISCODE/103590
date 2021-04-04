Attribute VB_Name = "mdlMenu"
Option Explicit

'*********************************************************************************************************************
'
'�˵���ش������
'
'*********************************************************************************************************************


'��ѯ��ݼ�����
Public Sub BindMenuShortcut(ByVal strProjectName As String, ByVal lngModule As Long, objMenu As Object)
    Dim strSQL As String
    Dim rsShoftcutCfg As ADODB.Recordset
    Dim objMain As Object

    strSQL = "select id,���Ƽ�,�ַ���,�����,�˵�ID from ��ݹ�����Ϣ where ��Ŀ=upper([1]) and ģ���=[2]"

    Set rsShoftcutCfg = zlDatabase.OpenSQLRecord(strSQL, "�󶨲˵���ݼ�", strProjectName, lngModule)
    
    Set objMain = objMenu
    
    Call RecursionBindMenu(objMain, objMenu.ActiveMenuBar, rsShoftcutCfg)
End Sub


'�󶨲˵���ݷ�ʽ(�ݹ���ð󶨿�ݲ˵�)
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

'�󶨵����˵��Ŀ�ݷ�ʽ
Private Sub BindMenuItemShortcut(cbrMain As Object, cbrControl As Object, rsShoftcutCfg As ADODB.Recordset)
    If rsShoftcutCfg Is Nothing Then Exit Sub
    
    Dim lngFuncKey As Long
    Dim lngCharKey As Long
    Dim lngCommandKey As Long
    
    Dim strKeyAlias As String

    rsShoftcutCfg.Filter = "�˵�ID=" & cbrControl.ID
    
    If rsShoftcutCfg.RecordCount > 0 Then
        lngFuncKey = Val(Nvl(rsShoftcutCfg!���Ƽ�))
        lngCharKey = Val(Nvl(rsShoftcutCfg!�ַ���))
        strKeyAlias = Nvl(rsShoftcutCfg!�����)

        'F8�̶�Ϊ��ݼ��ɼ�ʹ��
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
            
            '�󶨲˵���ݼ�
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
    
    
    'Begin----------------------�鿴�˵�--------------------------------------
    If Not (objViewMenu Is Nothing) Then
        Set cbrMenuBar = objViewMenu
        
        With cbrMenuBar.CommandBar
            Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(T)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 1
            
                With cbrControl.CommandBar '�����˵�
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(0)")
                        cbrPopControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                        
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(1)")
                        cbrPopControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                End With
            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(S)")
                cbrControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
        End With
    End If

    'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
    If Not (objHelpMenu Is Nothing) Then
        Set cbrMenuBar = objHelpMenu
        
        With cbrMenuBar.CommandBar
            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_Help, "��������(M)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 901
                
            Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_Help_Web, "WEB�ϵ�����(W)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 1
                
                With cbrControl.CommandBar
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Forum, "������̳(0)")
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                        
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Home, "������ҳ(1)")
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                        
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(2)")
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 9022
                End With
                
            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_About, "���ڡ�(A)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 1
        End With
    End If
End Sub

'*********************************************************************************************************************
