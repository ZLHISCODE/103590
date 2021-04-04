VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmMipClientShell 
   BorderStyle     =   0  'None
   ClientHeight    =   3120
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   6525
   Icon            =   "frmMipClientShell.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picNotify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2910
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   1425
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picService 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2925
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   225
   End
   Begin XtremeSuiteControls.PopupControl pct 
      Left            =   240
      Top             =   900
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
   End
   Begin VB.Image imgService 
      Height          =   240
      Index           =   0
      Left            =   3750
      Picture         =   "frmMipClientShell.frx":6852
      Top             =   285
      Width           =   240
   End
   Begin VB.Image imgNotify 
      Height          =   240
      Index           =   0
      Left            =   3420
      Picture         =   "frmMipClientShell.frx":7254
      Top             =   1455
      Width           =   240
   End
   Begin VB.Image imgService 
      Height          =   240
      Index           =   1
      Left            =   3390
      Picture         =   "frmMipClientShell.frx":8F4E
      Top             =   255
      Width           =   240
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   255
      Top             =   180
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmMipClientShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
Private mstrMessageKey As String
Private mstrMessageTopic As String
Private mstrMessageText As String
Private mbytMessageLinkType As Byte
Private mstrMessageLinkTitle As String
Private mstrMessageLinkPara As String
    
Private mclsMipReceiptData As clsMipReceiptData
Private mclsMipSystemData As clsMipSystemData
Private mblnNotifyIcon As Boolean
Private mobjXML As Object
Private mstrLogFile As String
Private mstrDataFile As String
Private mstrSysFile As String
Private mstrTitle As String
Private msglstart As Single

Public WithEvents mfrmMipComView As frmMipComView
Attribute mfrmMipComView.VB_VarHelpID = -1
Public WithEvents mfrmMipComOption As frmMipComOption
Attribute mfrmMipComOption.VB_VarHelpID = -1

Public Event OpenLink(ByVal bytLinkType As Byte, ByVal strLinkPara As String)

Public Event OptionChanged()

'######################################################################################################################
Public Function Initialize(ByVal strSysFile As String, ByVal strDataFile As String, ByVal strLogFile As String) As Boolean
    '******************************************************************************************************************
    '���ܣ���ʼ��
    '��������
    '���أ���ʼ���ɹ�����True,���򷵻�False
    '******************************************************************************************************************
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String
    Dim objFso As New FileSystemObject
    
    mblnNotifyIcon = False
    
    mstrSysFile = strSysFile
    mstrDataFile = strDataFile
    mstrLogFile = strLogFile
    mstrTitle = "������Ϣ����ƽ̨�ͻ���"
    
    Call InitCommandBar
    Call AddIcon(picService.hWnd, imgService(0).Picture, mstrTitle & "��δ���У�")
            
    Set mclsMipSystemData = New clsMipSystemData
    Set mclsMipReceiptData = New clsMipReceiptData
    
    If objFso.FolderExists(App.Path & "\Data") = False Then Call objFso.CreateFolder(App.Path & "\Data")
    
    strFile = App.Path & "\Data\zlMipClientShell_icon_111.ico"
    If objFso.FileExists(strFile) = False Then
        arrData = LoadResData(111, "CUSTOM")
        intFile = FreeFile
        Open strFile For Binary As intFile
        Put intFile, , arrData()
        Close intFile
    End If
    
    strFile = App.Path & "\Data\zlMipClientShell_icon_112.ico"
    If objFso.FileExists(strFile) = False Then
        arrData = LoadResData(112, "CUSTOM")
        intFile = FreeFile
        Open strFile For Binary As intFile
        Put intFile, , arrData()
        Close intFile
    End If
    
    strFile = App.Path & "\Data\zlMipClientShell_icon_113.ico"
    If objFso.FileExists(strFile) = False Then
        arrData = LoadResData(113, "CUSTOM")
        intFile = FreeFile
        Open strFile For Binary As intFile
        Put intFile, , arrData()
        Close intFile
    End If
    
    Initialize = mclsMipSystemData.Initialize(mstrSysFile) And mclsMipReceiptData.Initialize(mstrDataFile)
    Call mclsMipReceiptData.OpenDataFile
    
End Function

Private Function AddPopupControlItem(ByRef objPopupControl As Object, _
                                        ByRef objRect As RECT, _
                                        ByVal strText As String, _
                                        Optional blnFontBold As Boolean = False) As PopupControlItem
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objItem As PopupControlItem
    
    Set objItem = objPopupControl.AddItem(objRect.Left, objRect.Top, objRect.Right, objRect.Bottom, strText)
    objItem.Bold = blnFontBold
    
    Set AddPopupControlItem = objItem
    
End Function

Private Sub PopupMessage(ByVal strTitle As String, ByVal strInfo As String, ByVal strLink As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objItem As PopupControlItem
    Dim objRect As RECT
    Dim strText As String
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim strChar As String
    Dim strLine As String
    Dim intLine As Integer
    Dim lngWave As Long
    Dim lngAlert As Long
    Dim rsCondition As ADODB.Recordset
    Dim rsTmp As zlDataSQLite.SQLiteRecordset
    
    With pct
        .RemoveAllItems
                
        '����
        '--------------------------------------------------------------------------------------------------------------
        objRect.Left = 5
        objRect.Top = 13
        objRect.Right = objRect.Left + 16
        objRect.Bottom = objRect.Top + 16
        Set objItem = AddPopupControlItem(pct, objRect, strText)
        objItem.Id = 1
        Call objItem.SetIcon(LoadIcon("Data\zlMipClientShell_icon_112.ico", 16, 16), xtpPopupItemIconNormal)
        
        objRect.Left = objRect.Left + 20
        objRect.Top = objRect.Top + 2
        objRect.Right = 200
        objRect.Bottom = objRect.Top + 18
        
        Set objItem = AddPopupControlItem(pct, objRect, strTitle, True)
        objItem.Id = 2
        objItem.Hyperlink = False
                
        '--------------------------------------------------------------------------------------------------------------
        objRect.Left = 25
        objRect.Top = objRect.Bottom + 2
        objRect.Right = 290
        
        strText = strInfo
        intCount = Len(strText)
        For intLoop = 1 To intCount
            strChar = Mid(strText, intLoop, 1)
            If Me.TextWidth(strLine & strChar) > (objRect.Right - objRect.Left) * Screen.TwipsPerPixelX Or strChar = Chr(10) Then
                
                If strLine <> "" Then
                    intLine = intLine + 1
                    If intLine <= 6 Then
                        objRect.Top = objRect.Bottom + 1
                        objRect.Bottom = objRect.Top + 14
                        Set objItem = AddPopupControlItem(pct, objRect, strLine)
                        objItem.Id = 3
                        objItem.Hyperlink = False
                    Else
                        Exit For
                    End If
                    If strChar = Chr(10) Then
                        strLine = ""
                    Else
                        strLine = strChar
                    End If
                End If
            Else
                strLine = strLine & strChar
            End If
        Next
        
        If strLine <> "" Then
            intLine = intLine + 1
            If intLine <= 6 Then
                objRect.Top = objRect.Bottom + 1
                objRect.Bottom = objRect.Top + 14
                Set objItem = AddPopupControlItem(pct, objRect, strLine)
                objItem.Id = 3
                objItem.Hyperlink = False
            End If
        End If
                        
        '����
        '--------------------------------------------------------------------------------------------------------------
        If strLink <> "" Then
            objRect.Left = 5
            objRect.Top = 150 - 20
            objRect.Right = objRect.Left + 16
            objRect.Bottom = objRect.Top + 16
            Set objItem = AddPopupControlItem(pct, objRect, strText)
            objItem.Id = 4
            Call objItem.SetIcon(LoadIcon("Data\zlMipClientShell_icon_113.ico", 16, 16), xtpPopupItemIconNormal)
            
            objRect.Left = objRect.Left + 20
            objRect.Top = objRect.Top
            objRect.Right = 270
            objRect.Bottom = objRect.Top + 14
            strText = strLink
            Set objItem = AddPopupControlItem(pct, objRect, strText)
            objItem.Id = 5
            objItem.TextColor = RGB(0, 0, 255)
            objItem.TextAlignment = DT_LEFT Or DT_WORDBREAK
        End If
        
        '�رհ�ť
        '--------------------------------------------------------------------------------------------------------------
        objRect.Left = 274
        objRect.Top = 10
        objRect.Right = objRect.Left + 16
        objRect.Bottom = 26
        Set objItem = AddPopupControlItem(pct, objRect, "")
        Call objItem.SetIcon(LoadIcon("Data\zlMipClientShell_icon_111.ico", 16, 16), xtpPopupItemIconNormal)
        objItem.Id = 6
        objItem.Button = True
                
        .SetSize 300, 150
    
    End With
    
End Sub


Public Sub ShowMessage(ByVal strMessageContent As String)
    '******************************************************************************************************************
    '���ܣ���ʾ��Ϣ
    '������strMessageContent:XML��ʽ����Ϣ����
    '���أ���
    '******************************************************************************************************************
    Dim lngWave As Long
    Dim lngAlert As Long
    Dim intTransparency As Integer
    Dim rsCondition As ADODB.Recordset
    Dim rsTmp As zlDataSQLite.SQLiteRecordset
    
    mstrMessageKey = ""
    mstrMessageTopic = ""
    mstrMessageText = ""
    mbytMessageLinkType = 0
    mstrMessageLinkTitle = ""
    mstrMessageLinkPara = ""
    
    On Error GoTo errHand
    
    Do While pct.State <> xtpPopupStateClosed
        DoEvents
    Loop
        
    '������Ϣ
    '------------------------------------------------------------------------------------------------------------------
    Set mobjXML = InitXMLDoc
    Call mobjXML.loadXML(strMessageContent)

    mstrMessageTopic = ReadData("topic")
    mstrMessageText = ReadData("text")
    mbytMessageLinkType = Val(ReadData("link/type"))
    mstrMessageLinkTitle = ""
    mstrMessageLinkPara = ""
    If mbytMessageLinkType > 0 Then
        mstrMessageLinkTitle = ReadData("link/title")
        mstrMessageLinkPara = ReadData("link/para")
    End If
    If mstrMessageTopic = "" Then mstrMessageTopic = "������Ϣ"
    
    '��д��־
    '------------------------------------------------------------------------------------------------------------------
    If mclsMipReceiptData.OpenDataFile = True Then
        mstrMessageKey = mclsMipReceiptData.InsertReceiveMessage(mstrMessageText, mstrMessageTopic, mbytMessageLinkType, mstrMessageLinkTitle, mstrMessageLinkPara)
'        Call mclsMipReceiptData.CloseDataFile
    End If
        
    '��ʾ������Ϣͼ��
    '------------------------------------------------------------------------------------------------------------------
    If mblnNotifyIcon = False Then Call AddIcon(picNotify.hWnd, imgNotify(0).Picture, "��������Ϣ")
    
    lngWave = 0
    lngAlert = 5
    intTransparency = 5
    If mclsMipSystemData.OpenDataFile() = True Then
        Set rsCondition = CreateCondition

        '��Ϣ��������
        Call SetCondition(rsCondition, "�������", "1")
        rsTmp = mclsMipSystemData.GetPara("Filter", rsCondition)
        If rsTmp.DataSet.BOF = False Then
            lngWave = NVL(rsTmp.DataSet("Para_Value").Value)
        End If

        '��Ϣͣ��ʱ��
        Call SetCondition(rsCondition, "�������", "2")
        rsTmp = mclsMipSystemData.GetPara("Filter", rsCondition)
        If rsTmp.DataSet.BOF = False Then
            lngAlert = NVL(rsTmp.DataSet("Para_Value").Value)
        End If
        
        'ð�ݴ���͸����
        Call SetCondition(rsCondition, "�������", "5")
        rsTmp = mclsMipSystemData.GetPara("Filter", rsCondition)
        If rsTmp.DataSet.BOF = False Then
            intTransparency = NVL(rsTmp.DataSet("Para_Value").Value, "5")
        End If
        
        Call mclsMipSystemData.CloseDataFile
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    With pct
        .Animation = 2
        .AnimateDelay = 800
        .ShowDelay = lngAlert * 1000
        .Transparency = 255 - Int(255 * intTransparency / 20)
        .VisualTheme = xtpPopupThemeOffice2003
        
        '��������
        If lngWave > 0 Then Call PlayWave(lngWave)
        
        '��ʾ��Ϣ
        Call PopupMessage(mstrMessageTopic, mstrMessageText, mstrMessageLinkTitle)
        
        .Show
    End With
    
    Exit Sub
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
End Sub


Public Sub UpdateConnectState(ByVal blnConnectState As Boolean)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '˵����
    '******************************************************************************************************************
        
    If blnConnectState = False Then
        Call ModifyIcon(picService.hWnd, imgService(0).Picture, mstrTitle & "����ֹͣ��")
    Else
        Call ModifyIcon(picService.hWnd, imgService(1).Picture, mstrTitle & "�������У�")
    End If
    
End Sub

Private Function ReadData(ByVal strNode As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������strNode:Meta/Release
    '���أ�
    '******************************************************************************************************************
    Dim strData As String
    Dim objNode As Object
        
    strNode = ".//" & Replace(strNode, "/", "//")
    
    Set objNode = mobjXML.selectSingleNode(strNode)
    strData = objNode.Text
    
    ReadData = RestoreSpecialChar(strData)
End Function

Private Function RestoreSpecialChar(ByVal strXmlText As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strXML As String
    
    strXML = strXmlText
    If InStr(strXML, "&lt;") > 0 Then strXML = Replace(strXML, "&lt;", "<")
    If InStr(strXML, "&gt;") > 0 Then strXML = Replace(strXML, "&gt;", ">")
    If InStr(strXML, "&amp;") > 0 Then strXML = Replace(strXML, "&amp;", "&")
    If InStr(strXML, "&apos;") > 0 Then strXML = Replace(strXML, "&apos;", "'")
    If InStr(strXML, "&quot;") > 0 Then strXML = Replace(strXML, "&quot;", """")
    
    RestoreSpecialChar = strXML
End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call CommandBarInit(cbsMain)
    Set cbsMain.Icons = frmMipResource.imgPublic.Icons
    cbsMain.Options.LargeIcons = True
    
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
        
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
    Case conMenu_File_Parameter
        
        
        If mfrmMipComOption Is Nothing Then Set mfrmMipComOption = New frmMipComOption
        
        If mfrmMipComOption.ShowDialog(Me, mstrSysFile) Then
            '�в����仯
'            RaiseEvent OptionChanged
        End If
    Case conMenu_View_ShowHistory
        If mfrmMipComView Is Nothing Then Set mfrmMipComView = New frmMipComView
        If Not (mfrmMipComView Is Nothing) Then Call mfrmMipComView.ShowForm(Me, mstrDataFile)
    Case conMenu_View_Jump
        Call frmMipClientRunlog.ShowForm(Me, mstrLogFile)
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim frmThis As Form
    
    On Error Resume Next
    
    If Not (mfrmMipComOption Is Nothing) Then
        Unload mfrmMipComOption
        Set mfrmMipComOption = Nothing
    End If
    
    If Not (mfrmMipComView Is Nothing) Then
        Unload mfrmMipComView
        Set mfrmMipComView = Nothing
    End If
    
    '�رձ���������
    For Each frmThis In Forms
        If frmThis.Caption <> Me.Caption Then Unload frmThis
    Next

    Set mclsMipSystemData = Nothing
    
    Call RemoveIcon(picNotify.hWnd)
    Call RemoveIcon(picService.hWnd)
End Sub

Public Function ShowConetneMenu(Optional ByVal bytPlace As Byte = 1) As CommandBar
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrPopupItem2 As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    
    '�����˵�����
    
    On Error GoTo errHand
    
    Set cbrPopupBar = cbsMain.Add("�����˵�", xtpBarPopup)
    
    Select Case bytPlace
    Case 1
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_File_Parameter, "ѡ������(&O)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_ShowHistory, "��Ϣ����(&M)")
        cbrPopupItem.BeginGroup = True
                
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_Jump, "������־(&R)")
        
    End Select
    
    Set ShowConetneMenu = cbrPopupBar
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
'    If zlComLib.ErrCenter = 1 Then
'        Resume
'    End If
End Function

Private Sub SetReadedFlag()
    '���Ķ���ǰ��Ϣ
    If mstrMessageKey <> "" Then
        If mclsMipReceiptData.OpenDataFile = True Then
            Call mclsMipReceiptData.UpdateReceiveMessageReaded(mstrMessageKey)
'            Call mclsMipReceiptData.CloseDataFile
        End If
    End If
End Sub

Private Sub mfrmMipComOption_OptionChanged()
    RaiseEvent OptionChanged
End Sub

Private Sub mfrmMipComView_AfterReadMessage()
    Call CheckMessageIcon
End Sub

Private Sub mfrmMipComView_OpenLink(ByVal bytLinkType As Byte, ByVal strLinkPara As String)
    If bytLinkType > 0 Then
        
        If IsWindowEnabled(glngParentForm) = 0 Then
            Screen.MousePointer = 0
            MsgBox "��ǰϵͳ�����ж�ռ��ģ̬����򿪣����ȹر���ִ�е�ǰ������", vbInformation, gstrSysName
            Exit Sub
        End If
        
        RaiseEvent OpenLink(bytLinkType, strLinkPara)
                    
    End If
End Sub

Private Sub pct_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    '
    '�����Ķ���־
    Call SetReadedFlag
    Call CheckMessageIcon
    
    '�ֱ�ִ��
    Select Case Item.Id
    Case 5      '����
                
        If mbytMessageLinkType > 0 Then
            
            If IsWindowEnabled(glngParentForm) = 0 Then
                Screen.MousePointer = 0
                MsgBox "��ǰϵͳ�����ж�ռ��ģ̬����򿪣����ȹر���ִ�е�ǰ������", vbInformation, gstrSysName
                Exit Sub
            End If
            
            RaiseEvent OpenLink(mbytMessageLinkType, mstrMessageLinkPara)
                        
        End If
    Case 6          '�ر�
        pct.Close
    End Select
End Sub

Private Sub pct_StateChanged()
    Dim blnCheck As Boolean
    Dim sglEnd As Single
    
    Select Case pct.State
    '------------------------------------------------------------------------------------------------------------------
    Case xtpPopupStateShow
        msglstart = Timer
    '------------------------------------------------------------------------------------------------------------------
    Case xtpPopupStateClosed
        msglstart = 0
        Call CheckMessageIcon
    '------------------------------------------------------------------------------------------------------------------
    Case xtpPopupStateCollapsing
        '
        If msglstart > 0 Then

            sglEnd = Timer
            
            '����ͣ��ʱ���ж�Ϊ���Ķ������߳���Ҳ��Ϊ���Ķ�
            On Error Resume Next
            Err = 0
            If (sglEnd - msglstart) * 1000 > pct.ShowDelay + 10 Then blnCheck = True
            If Err <> 0 Then blnCheck = True
            On Error GoTo 0
            
            msglstart = 0
            
            If blnCheck Then
                Call SetReadedFlag
                Call CheckMessageIcon
            End If
            
        End If
        
    End Select
    
End Sub

Private Sub CheckMessageIcon()
    If mclsMipReceiptData.OpenDataFile = True Then
        If mclsMipReceiptData.ExistUnReadMessge = False Then
            Call RemoveIcon(picNotify.hWnd)
            mblnNotifyIcon = False
        End If
    End If
End Sub

Private Sub picNotify_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '--------------------------------------------------------------------------------------------------
    '����:  ����picNotify�ĸ��ִ����¼�,��Ҫ�����Զ�������ع���(�����д)
    '--------------------------------------------------------------------------------------------------

    Select Case Hex(X) '
        Case "1E3C"     'Right-Button-Down
        Case "1E4B"     'Right-Button-Up
        Case "1830"     'Right-Button-Down LARGE FONTS '
        Case "1E1E"     'Left-Button-up
        Case "1E0F"     'Left-Button-Down '
        Case "1E2D"     'Left-Button-Double-Click '
            If mfrmMipComView Is Nothing Then Set mfrmMipComView = New frmMipComView
            If Not (mfrmMipComView Is Nothing) Then Call mfrmMipComView.ShowForm(Me, mstrDataFile, True)
'            Call frmMipComView.ShowForm(Me, mstrDataFile, True)
        Case "1824"     'Left-Button-Double-Click LARGE FONTS
        Case "1E5A"     'Right-Button-Double-Click '
    End Select '

End Sub

Private Sub picService_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '--------------------------------------------------------------------------------------------------
    '����:  ����picNotify�ĸ��ִ����¼�,��Ҫ�����Զ�������ع���(�����д)
    '--------------------------------------------------------------------------------------------------

    Select Case Hex(X) '
        Case "1E3C"     'Right-Button-Down
        Case "1E4B"     'Right-Button-Up
            Call ShowConetneMenu(1).ShowPopup
        Case "1830"     'Right-Button-Down LARGE FONTS '
        Case "1E1E"     'Left-Button-up
        Case "1E0F"     'Left-Button-Down '
        Case "1E2D"     'Left-Button-Double-Click '
            '
        Case "1824"     'Left-Button-Double-Click LARGE FONTS
        Case "1E5A"     'Right-Button-Double-Click '
    End Select '
End Sub
