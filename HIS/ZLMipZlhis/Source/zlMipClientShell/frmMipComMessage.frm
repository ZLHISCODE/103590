VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmMspClientShell 
   BorderStyle     =   0  'None
   ClientHeight    =   3120
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   6525
   Icon            =   "frmMspComMessage.frx":0000
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
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picService 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   3285
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   1395
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Timer tmrModule 
      Index           =   0
      Left            =   315
      Top             =   1515
   End
   Begin VB.Image imgNotify 
      Height          =   240
      Index           =   0
      Left            =   330
      Picture         =   "frmMspComMessage.frx":6852
      Top             =   495
      Width           =   240
   End
   Begin VB.Image imgService 
      Height          =   240
      Index           =   0
      Left            =   3495
      Picture         =   "frmMspComMessage.frx":854C
      Top             =   450
      Width           =   240
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   1305
      Top             =   435
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmMspClientShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
Private mstrMessageKey As String
Private mstrMessageTopic As String
Private mstrMessageText As String
Private mstrMessageLinkType As String
Private mstrMessageLinkTitle As String
Private mstrMessageLinkPara As String
    
Private mclsMspData As clsMspData
Private WithEvents mfrmMspComAlert As frmMspComAlert
Attribute mfrmMspComAlert.VB_VarHelpID = -1
Private mblnNotifyIcon As Boolean
Private mobjXML As DOMDocument

Public Event OpenLink(ByVal bytLinkType As Byte, ByVal strLinkPara As String)

'######################################################################################################################
Public Function Initialize() As Boolean
    '******************************************************************************************************************
    '���ܣ���ʼ��
    '��������
    '���أ���ʼ���ɹ�����True,���򷵻�False
    '******************************************************************************************************************
    mblnNotifyIcon = False
    
    Call InitCommandBar
    Call AddIcon(picService.hWnd, imgService(0).Picture, "��Ϣ����ƽ̨�ͻ��˷���")
            
    Initialize = True
    
End Function

Public Sub ShowMessage(ByVal strMessageContent As String)
    '******************************************************************************************************************
    '���ܣ���ʾ��Ϣ
    '������strMessageContent:XML��ʽ����Ϣ����
    '���أ���
    '******************************************************************************************************************
    
    mstrMessageKey = ""
    mstrMessageTopic = ""
    mstrMessageText = ""
    mstrMessageLinkType = ""
    mstrMessageLinkTitle = ""
    mstrMessageLinkPara = ""
    
    If mfrmMspComAlert Is Nothing Then Set mfrmMspComAlert = New frmMspComAlert
    If Not (mfrmMspComAlert Is Nothing) Then
        Set mobjXML = New DOMDocument
        Call mobjXML.loadXML(strMessageContent)
        
        mstrMessageTopic = ReadData("topic")
        mstrMessageText = ReadData("text")
        mstrMessageLinkType = ReadData("link/type")
        
        Select Case mstrMessageLinkType
        Case "����", "ģ��"
            mstrMessageLinkTitle = ReadData("link/title")
            mstrMessageLinkPara = ReadData("link/para")
        End Select
    
        Call mfrmMspComAlert.ShowAlert(mstrMessageTopic, mstrMessageText, ReadData("link/type"), mstrMessageLinkTitle, mstrMessageLinkPara)
    End If
        
End Sub

Private Function ReadData(ByVal strNode As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������strNode:Meta/Release
    '���أ�
    '******************************************************************************************************************
    Dim strData As String
    Dim objNode As IXMLDOMNode
        
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

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
        
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
    Case 1
        Call frmMspComOption.ShowDialog(Me)
    Case 2
        frmMspComView.Show , Me
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim frmThis As Form
    
    On Error Resume Next
    
    '�رձ���������
    For Each frmThis In Forms
        If frmThis.Caption <> Me.Caption Then Unload frmThis
    Next
            
    If Not (mfrmMspComAlert Is Nothing) Then
        Unload mfrmMspComAlert
        Set mfrmMspComAlert = Nothing
    End If
    
    Call RemoveIcon(picService.hWnd)
End Sub

Private Sub mfrmMspComAlert_AfterShowMessage()
    '����Ƿ�����Ϣû�в��ģ�������������ʾͼ�꣬���û�У�������ͼ��
        
    Set mclsMspData = New clsMspData
    Call mclsMspData.Initialize(App.Path & "\Data\zlMspComMessage.db")
        
    If mclsMspData.IsUnReadMessge = True Then
        Call RemoveIcon(picNotify.hWnd)
        mblnNotifyIcon = False
    End If
    
End Sub

Private Sub mfrmMspComAlert_BeforeShowMessage()


    '��д��־
    mstrMessageKey = mclsMspData.InsertReceiveMessage(mstrMessageText, mstrMessageText, mstrMessageLinkType, mstrMessageLinkTitle, mstrMessageLinkPara)
        
    '��ʾ������Ϣͼ��
    If mblnNotifyIcon = False Then
        Call AddIcon(picNotify.hWnd, imgNotify(0).Picture, "��������Ϣ")
    End If
End Sub

Private Sub mfrmMspComAlert_OpenLink(ByVal bytLinkType As Byte, ByVal strLinkPara As String)
        
    If strLinkPara <> "" Then
        
        If IsWindowEnabled(glngParentForm) = 0 Then
            Screen.MousePointer = 0
            MsgBox "��ǰϵͳ�����ж�ռ��ģ̬����򿪣����ȹر���ִ�е�ǰ������", vbInformation, gstrSysName
            Exit Sub
        End If
                                
        RaiseEvent OpenLink(bytLinkType, strLinkPara)
    
    End If
    
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
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, 1, "ѡ������(&O)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, 2, "ͨ����Ϣ(&M)")
        cbrPopupItem.BeginGroup = True
        cbrPopupItem.DefaultItem = True
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, 3, "������־(&R)")
        
    End Select
    
    Set ShowConetneMenu = cbrPopupBar
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
'    If zlComLib.ErrCenter = 1 Then
'        Resume
'    End If
End Function

Private Sub mfrmMspComAlert_ReadMessage()
    '���Ķ���ǰ��Ϣ
    If mstrMessageKey <> "" Then
        Set mclsMspData = New clsMspData
        Call mclsMspData.Initialize(App.Path & "\Data\zlMspComMessage.db")
        Call mclsMspData.UpdateReceiveMessageReaded(mstrMessageKey)
        Call mclsMspData.CloseDataFile
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
            Call frmMspComView.ShowForm(Me)
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
