VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMain 
   Caption         =   "�Զ����ѷ���"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8175
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   8175
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picNotify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   5100
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   4110
      Visible         =   0   'False
      Width           =   225
   End
   Begin MSWinsockLib.Winsock winSock 
      Left            =   3915
      Top             =   4095
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "FRCHEN"
   End
   Begin VB.Timer tmrMessage 
      Interval        =   1
      Left            =   225
      Top             =   4005
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   4830
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":038A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8017
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "2017/7/28"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "13:21"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2685
      Left            =   210
      TabIndex        =   1
      Top             =   1125
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   4736
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����վ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�˿ں�"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�û�����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "���ݿ��û�"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "IP��ַ"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   8175
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   315
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "�������ѷ���"
               Object.Tag             =   "����"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ֹͣ"
               Key             =   "ֹͣ"
               Object.ToolTipText     =   "ֹͣ���ѷ���"
               Object.Tag             =   "ֹͣ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�����"
               Object.Tag             =   "�˳�"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   7485
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   6825
      Top             =   705
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":391C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4096
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   6210
      Top             =   705
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":44CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B38
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D52
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgNotify 
      Height          =   240
      Left            =   5295
      Picture         =   "frmMain.frx":5F6C
      Top             =   90
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileStart 
         Caption         =   "��������(&S)"
      End
      Begin VB.Menu mnuFileStop 
         Caption         =   "ֹͣ����(&D)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileParm 
         Caption         =   "��������(&P)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLogin 
         Caption         =   "���µ�¼(&L)"
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileHide 
         Caption         =   "�������ѷ���(&H)"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�����(&X)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "��������(&T)"
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnStartUp As Boolean
Private mlngCount As Long
Private mlngPort As Long
Private mstrLocalIP As String
Private mblnCancel As Boolean
Private mblnTest As Boolean

Private Sub cmdFunc_Click(Index As Integer)
    Select Case Index
    Case 0
        If tbrThis.Buttons("ע��").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("ע��"))
    Case 1
        If tbrThis.Buttons("����").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("����"))
    Case 2
        If tbrThis.Buttons("ֹͣ").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("ֹͣ"))
    Case 3
        If tbrThis.Buttons("�˳�").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("�˳�"))
    Case 4
        If tbrThis.Buttons("�˿�").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("�˿�"))
    Case 5
        '
    End Select
End Sub

Private Sub Form_Activate()
    Dim rs As New ADODB.Recordset
    Dim varParam As Variant
    Dim strSQL As String
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
                                        
    
    Me.Caption = Me.Caption & " - [" & gstrUserName & IIf(gstrServer = "", "", "@" & gstrServer) & "]"
    
    gstrSysName = gstrProductName & "���"
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrSysName"), gstrSysName
    
    Call ApplyOEM(stbThis)
    Call ApplyOEM_Picture(Me, "Icon")
    
   
    '��ʽ:������;�˿ں�;״̬
    strSQL = "SELECT ����ֵ FROM zloptions WHERE ������=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, 7)
    
    If rs.BOF = False Then
        varParam = Split(zlCommFun.NVL(rs("����ֵ").Value, ""), ";")
        If UBound(varParam) < 2 Then
            mlngPort = 9999
            mstrLocalIP = winSock.LocalIP
        Else
            mlngPort = Val(varParam(1))            'ȡ���������õĶ˿ں�
            mstrLocalIP = Trim(varParam(0))
        End If
    Else
        mlngPort = 9999
        mstrLocalIP = winSock.LocalIP
    End If

    '��������
    Call tbrThis_ButtonClick(tbrThis.Buttons("����"))
        
    DoEvents
    
    Call AddIcon(picNotify.hWnd, imgNotify.Picture, Me.Caption)
    
    If rs.State = adStateOpen Then rs.Close
    
    Call mnuFileHide_Click
    
End Sub

Private Sub Form_Load()
    
    tmrMessage.Interval = 1
    tmrMessage.Enabled = False
    mblnStartUp = True
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Select Case Me.WindowState
    Case 1
        mnuFileHide.Caption = "��ʾ���ѷ���(&O)"
        Me.Hide
        Exit Sub
    End Select
    
    With lvw
        .Left = 0
        .Top = cbrThis.Height
        .Height = Me.ScaleHeight - .Top - stbThis.Height
        .Width = Me.ScaleWidth - .Left
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    mblnCancel = False
    
    If MsgBox("���Ƿ����Ҫ�˳��Զ����ѷ���", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
        Cancel = True
        mblnCancel = True
        Exit Sub
    End If

    On Error Resume Next
    If gcnOracle.State <> adStateClosed Then gcnOracle.Close
    Call RemoveIcon(picNotify.hWnd)
    
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lvw.SortKey = ColumnHeader.Index - 1 Then
        lvw.SortOrder = IIf(lvw.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvw.SortKey = ColumnHeader.Index - 1
        lvw.SortOrder = lvwAscending
    End If
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileHide_Click()
    If mnuFileHide.Caption = "�������ѷ���(&H)" Then
        Me.WindowState = 1
        Me.Hide
        mnuFileHide.Caption = "��ʾ���ѷ���(&O)"
    Else
        Me.WindowState = 0
        Me.Show
        mnuFileHide.Caption = "�������ѷ���(&H)"
    End If
End Sub

Private Sub mnuFileLogin_Click()
    'If MsgBox("���Ƿ����Ҫע����Ϣ���ѷ���", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Sub
    
    Unload Me
    If mblnCancel = False Then Call Main
End Sub

Private Sub mnuFileParm_Click()
    mnuFileHide.Enabled = False
    
    If frmParam.ShowEdit(Me, mlngPort, mstrLocalIP, IIf(tbrThis.Buttons("����").Enabled, 0, 1)) Then
            
    End If
    mnuFileHide.Enabled = True
End Sub


Private Sub mnuFileStart_Click()
    Select Case StartServer(mlngPort, mstrLocalIP)
    Case 0
        Call AdjustEnabledState(1)
    Case 1
        stbThis.Panels(2).Text = "���󣺶˿ڵ�ַ����ʹ�ã�"
    End Select
End Sub

Private Sub mnuFileStop_Click()
    winSock.Close
    tmrMessage.Enabled = False
    stbThis.Panels(2).Text = "���ѷ�����ֹͣ��"
    lvw.ListItems.Clear

    Call AdjustEnabledState(2)
End Sub


Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
    Shell "hh.exe  zl9SvrNotice.chm", vbNormalFocus
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub picNotify_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    '--------------------------------------------------------------------------------------------------
    '����:  ����picNotify�ĸ��ִ����¼�,��Ҫ�����Զ�������ع���(�����д)
    '--------------------------------------------------------------------------------------------------

    Select Case Hex(x) '
        Case "1E3C"     'Right-Button-Down
        Case "1E4B"     'Right-Button-Up
            Me.PopupMenu mnuFile
        Case "1830"     'Right-Button-Down LARGE FONTS '
        Case "1E1E"     'Left-Button-up
        Case "1E0F"     'Left-Button-Down '
        Case "1E2D"     'Left-Button-Double-Click '
            If mnuFileHide.Enabled Then Call mnuFileHide_Click
        Case "1824"     'Left-Button-Double-Click LARGE FONTS
        Case "1E5A"     'Right-Button-Double-Click '
    End Select '

End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "����"
        Call mnuFileStart_Click
    Case "ֹͣ"
        Call mnuFileStop_Click
    Case "����"
        Call mnuFileParm_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tmrMessage_Timer()
    
    If tmrMessage.Interval = 1 Then tmrMessage.Interval = 60000
    'Call ShowAlert(Me)
    
    tmrMessage.Enabled = False
    Call CheckNoticeAll
    tmrMessage.Enabled = True
    
End Sub

Private Sub winSock_DataArrival(ByVal bytesTotal As Long)
    Dim varItem As Variant
    Dim strData As String
    Dim varData As Variant
    Dim lngLoop As Long
    Dim strSQL As String
    Dim objItem As ListItem
    Dim strTmp As String
    
    On Error Resume Next
    
    winSock.GetData strData
    
    lngLoop = InStr(strData, "]")
    If lngLoop = 0 Then Exit Sub
    
    strTmp = Mid(strData, 2, lngLoop - 2)
    strData = Mid(strData, lngLoop + 1)
    
    Select Case strTmp
    Case "SYS-COMPUTER"             '����վ�ش��й��������
        
        '��ʽ:������;�˺ź�;�û���;�û�����
        
        varData = Split(strData, ";")
        
        If UCase(varData(5)) = UCase(mstrLocalIP) Then
            
            For lngLoop = lvw.ListItems.Count To 1 Step -1
                
                Set objItem = lvw.ListItems(lngLoop)
                If UCase(objItem.Text) = UCase(varData(0)) And _
                    UCase(objItem.SubItems(1)) = UCase(varData(1)) And _
                    UCase(objItem.SubItems(2)) = UCase(varData(2)) And _
                    UCase(objItem.SubItems(3)) = UCase(varData(3)) Then
                    
                    lvw.ListItems.Remove lngLoop
                    Exit For
                    
                End If
            Next
            
            Set objItem = lvw.ListItems.Add(, , varData(0), 1, 1)
            objItem.SubItems(1) = varData(1)
            objItem.SubItems(2) = varData(2)
            objItem.SubItems(3) = varData(3)
            objItem.SubItems(4) = GetClientIP(objItem.Text)
            objItem.ListSubItems(1).Tag = Val(varData(4))
            
        End If
        
    Case "SYS-DISCONNECT"           '����վ������������Ϣ
    
        '��ʽ:������;�˺ź�;�û���;�û�����
        
        varData = Split(strData, ";")
        If UCase(varData(5)) = UCase(mstrLocalIP) Then
            For lngLoop = lvw.ListItems.Count To 1 Step -1
                
                Set objItem = lvw.ListItems(lngLoop)
                If UCase(objItem.Text) = UCase(varData(0)) And _
                    UCase(objItem.SubItems(1)) = UCase(varData(1)) And _
                    UCase(objItem.SubItems(2)) = UCase(varData(2)) And _
                    UCase(objItem.SubItems(3)) = UCase(varData(3)) Then
                    
                    lvw.ListItems.Remove lngLoop
                    Exit For
                    
                End If
            Next
        End If
        
    Case "SYS-STARTUP"              '����վ����ʱ�����������������������Ϣ
    
        '��ʽ:������;�˺ź�;�û���;��������
        
        varData = Split(strData, ";")
        If UCase(varData(5)) = UCase(mstrLocalIP) Then
            For lngLoop = lvw.ListItems.Count To 1 Step -1
                
                Set objItem = lvw.ListItems(lngLoop)
                If UCase(objItem.Text) = UCase(varData(0)) And _
                    UCase(objItem.SubItems(1)) = UCase(varData(1)) And _
                    UCase(objItem.SubItems(2)) = UCase(varData(2)) And _
                    UCase(objItem.SubItems(3)) = UCase(varData(3)) Then
                            
                    '����վ���������������
                    tmrMessage.Enabled = False
                    Call CheckNoticeOne(objItem, True)
                    tmrMessage.Enabled = True
                    
                    Exit For
                    
                End If
            Next
        End If
    Case "SYS-READED"               '����վ�����Ļ�����Ϣ�Ѷ���־
    
        '��ʽ:�������;�û���
        
        varData = Split(strData, ";")
        
        If UCase(varData(5)) = UCase(mstrLocalIP) Then
            strSQL = "Zl_Zlnoticerec_Edit(1," & Val(varData(0)) & ",'" & varData(1) & "',Null,Null,Null,1,Null)"
            On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    Case "SYS-TEST"
        mblnTest = True
    End Select
    
    Exit Sub
    
errHand:
'    gcnOracle.RollbackTrans
End Sub

Private Function StartServer(Optional ByVal Port As Long = 1024, Optional ByVal LocalIP As String = "") As Long
    '--------------------------------------------------------------------------------------------------------
    '����:�����Զ����ѷ���
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    '��ʼ��
    winSock.Close
    winSock.Protocol = sckUDPProtocol
    
    stbThis.Panels(2).Text = "���������Զ����ѷ���...."
    DoEvents
    
    '��ʼ����
    Err = 0
    On Error Resume Next
    
'    winSock.Bind Port
    winSock.LocalPort = Port
    winSock.SendData ""
    
    If Err <> 0 Then
        MsgBox Err.Description, , gstrSysName
        
        stbThis.Panels(2).Text = ""
        StartServer = 1
                
        Exit Function
    End If
    On Error GoTo 0
    
    stbThis.Panels(2).Text = "���ѷ���������(IP:" & LocalIP & " Port:" & Port & ")��"
    
    StartServer = 0
    
End Function

Private Sub AdjustEnabledState(ByVal bytMode As Byte)
    If bytMode = 1 Then
        '�Ѿ�����
        tmrMessage.Enabled = True
        tbrThis.Buttons("����").Enabled = False
        tbrThis.Buttons("ֹͣ").Enabled = True
        
        mnuFileStart.Enabled = False
        mnuFileStop.Enabled = True
    Else
        '�Ѿ�ֹͣ
        tmrMessage.Enabled = False
        mnuFileStart.Enabled = True
        mnuFileStop.Enabled = False
    End If
    
    
    tbrThis.Buttons("ֹͣ").Enabled = mnuFileStop.Enabled
    tbrThis.Buttons("����").Enabled = mnuFileStart.Enabled
End Sub

Public Function UpdateRefresh(ByVal lngNewPort As Long, ByVal strLocalIP As String)
    '��������������������ֹͣ
    If tbrThis.Buttons("����").Enabled = False Then
        winSock.Close
        tmrMessage.Enabled = False
    End If
    
    '����¶˿ں��Ƿ���Ч
    If StartServer(lngNewPort, strLocalIP) <> 0 Then
        MsgBox "���õĶ˿ں���Ч���ͻ��", vbOKOnly, gstrSysName
        
        '���ԭ����������״̬�������Ҫ��������
        If tbrThis.Buttons("����").Enabled = False Then
            Call tbrThis_ButtonClick(tbrThis.Buttons("����"))
        End If
        
        Exit Function
    End If
    
    '���ԭ����������״̬����Ҫ���¶˿ں���������
    If tbrThis.Buttons("����").Enabled = False Then
        Call AdjustEnabledState(1)
    End If
    
    mlngPort = lngNewPort
    
    UpdateRefresh = True
    
End Function

Private Function CheckNoticeAll() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:ִ�����Ѽ��,����������
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strMachine As String
    
    For lngLoop = 1 To lvw.ListItems.Count
        lvw.ListItems(lngLoop).ListSubItems(2).Tag = ""
    Next
    
    strSQL = "SELECT DISTINCT USERNAME,TERMINAL FROM GV$Session WHERE USERNAME IS NOT NULL"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rs.BOF = False Then
        Do While Not rs.EOF
            strMachine = Trim(NVL(rs("TERMINAL").Value))
            If InStr(strMachine, "\") > 0 Then strMachine = Mid(strMachine, InStr(strMachine, "\") + 1)
            
            If strMachine <> "" Then

                If Chr(0) = Right(strMachine, 1) Then
                    strMachine = Mid(strMachine, 1, Len(strMachine) - 1)
                End If
                
                For lngLoop = 1 To lvw.ListItems.Count
                    If lvw.ListItems(lngLoop).ListSubItems(2).Tag = "" Then
                        If UCase(lvw.ListItems(lngLoop).Text) = UCase(strMachine) And _
                            UCase(lvw.ListItems(lngLoop).SubItems(3)) = UCase(NVL(rs("USERNAME").Value)) Then
                            
                            lvw.ListItems(lngLoop).ListSubItems(2).Tag = "CHECKED"
                            Call CheckNoticeOne(lvw.ListItems(lngLoop))
                            
                        End If
                    End If
                Next
            End If
            rs.MoveNext
        Loop
    End If
    
    'ɾ���Ѿ��쳣��ֹ�Ĺ���վ
    For lngLoop = lvw.ListItems.Count To 1 Step -1
        If lvw.ListItems(lngLoop).ListSubItems(2).Tag = "" Then
            lvw.ListItems.Remove lngLoop
        End If
    Next
    
End Function

Private Function CheckNoticeOne(ByVal Item As MSComctlLib.ListItem, Optional ByVal StartUp As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:���ָ���û���������Ϣ,����������
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim lngLoop As Long
    Dim str�������� As String
    Dim strSyss As String
    Dim strDepts As String
    
    mblnTest = False
    Item.Tag = ""
    
    '1.�������Ƿ񻹴��ڵ�¼״̬(��Ϊ����վ�쳣�˳�ʱ)
    
    '2.�ҳ��û�������ϵͳ����������
    strSQL = "Select ���, ����, �����, ������, ��װ����, ������װ, �汾�� From zlSystems "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rs.BOF Then Exit Function
    
    strSyss = "0"
    strDepts = "''"
    
    Do While Not rs.EOF
    
        On Error Resume Next
        
        Err = 0
        strSQL = "Select R.����id " & _
                        " From " & rs("������").Value & ".�ϻ���Ա�� U," & rs("������").Value & ".��Ա�� P," & rs("������").Value & ".������Ա R" & _
                        " Where U.��ԱID = P.ID And P.ID=R.��ԱID and U.�û���=[1] And (P.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or P.����ʱ�� Is Null) and R.ȱʡ=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Item.SubItems(3))
        If Err = 0 Then
            If rsTmp.BOF = False Then
                strSyss = strSyss & "," & rs("���").Value
                
                strSQL = "SELECT ���� FROM " & rs("������").Value & ".���ű� START WITH ID=[1] CONNECT BY PRIOR �ϼ�id=ID"
                Set rs2 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rsTmp("����id").Value)
                If rs2.BOF = False Then
                    Do While Not rs2.EOF
                        strDepts = strDepts & ",'" & UCase(rs2("����").Value) & "'"
                        rs2.MoveNext
                    Loop
                End If
            End If
        End If
        
        On Error GoTo 0
        
        rs.MoveNext
    Loop
        
    '3.�ҳ����û���������Ϣ
    strSQL = "SELECT A.���, A.ϵͳ, A.��������, A.��������, A.���ѱ���, A.��������, A.���Ѵ���, A.����˳��, A.�������, A.��������, A.��ʼʱ��, A.��ֹʱ��, B.����ID AS ģ��,B.���� AS ��������,B.ϵͳ As ����ϵͳ " & _
                "FROM zlNotices A," & _
                    "(SELECT Id , ���, ����, ˵��, ����, ��ӡ��, ��ֽ, Ʊ��, ��ӡ��ʽ, ϵͳ, ����id, ����, �޸�ʱ��, ����ʱ��, ��ֹ��ʼʱ��, ��ֹ����ʱ��, ִ�п�ʼʱ��, ִ�н���ʱ��, ִ����Ա, ���ִ��ʱ�� FROM zlReports WHERE ����ʱ�� IS NOT NULL) B " & _
                "WHERE (A.��� IN (" & _
                                    "SELECT ������� FROM zlNoticeUsr WHERE ���Ѷ���=0 " & _
                                    "Union " & _
                                    "SELECT ������� FROM zlNoticeUsr WHERE ���Ѷ���=1 AND UPPER(��������)=[1] " & _
                                    "Union " & _
                                    "SELECT ������� FROM zlNoticeUsr WHERE ���Ѷ���=2 AND UPPER(��������) IN (" & strDepts & ") " & _
                                    "Union " & _
                                    "SELECT ������� FROM zlNoticeUsr WHERE ���Ѷ���=3 AND UPPER(��������)=[2] " & _
                                ") " & _
                        "OR NOT EXISTS (SELECT ������� FROM zlNoticeUsr C WHERE C.�������=A.���))" & _
                    "AND B.���(+) = A.���ѱ��� " & _
                    "AND (A.ϵͳ IN (" & strSyss & ") OR A.ϵͳ IS NULL) " & _
                    "AND A.��ʼʱ�� <= SYSDATE And (A.��ֹʱ�� >= SYSDATE Or A.��ֹʱ�� Is Null) " & _
                    "AND " & IIf(StartUp = False, " A.������� IS NOT NULL", " A.������� IS NULL")
                    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(Item.SubItems(2)), UCase(Item.Text))
    If rs.BOF = False Then
    
        Item.Tag = ""
        
        Do While Not rs.EOF
            
            If CheckNotice(Item.SubItems(2), rs("���").Value, str��������) Then
                'Ҫ����,����������Ϣ
                                
                '--����û��Ƿ��Ѿ�����-----------------------------------------------------------------------------------------------------------
                strSQL = "SELECT �������, �û���, ���ʱ��, �����, ���ѱ�־, ����ʱ��, �Ѷ���־, �������� FROM zlNoticeRec WHERE �������=[1] AND �û���=[2] AND �Ѷ���־<>1"
                
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rs("���").Value, Item.SubItems(2))
                If rsTmp.BOF = False Then

                    '���ѹ�,�ټ���Ƿ�����Ϣ
                    If rsTmp("���ѱ�־").Value > 0 Then
                        Item.Tag = Item.Tag & "[INFOITEM-BEGIN']" & NVL(rsTmp("��������").Value, 0) & "[''']" & NVL(rs("��������").Value, 0) & "[''']" & NVL(rs("��������").Value, "") & "[''']" & NVL(rs("ϵͳ").Value, 0) & "[''']" & NVL(rs("ģ��").Value, 0) & "[''']" & winSock.LocalHostName & "[''']" & NVL(rs("���Ѵ���").Value, 0) & "[''']" & NVL(rs("����ϵͳ").Value, 0)
                    End If
                    
                End If
                '--����û��Ƿ��Ѿ�����-----------------------------------------------------------------------------------------------------------
                
            
            End If
            
            rs.MoveNext
        Loop
        
        '��������
        If Item.Tag <> "" Then
            If Left(Item.Tag, 17) = "[INFOITEM-BEGIN']" Then Item.Tag = Mid(Item.Tag, 18)
            Call SendMessage(Item.SubItems(4), Val(Item.SubItems(1)), Item.SubItems(2), Item.Tag)
        End If
    End If
End Function

Private Function CheckNotice(ByVal strUser As String, ByVal lngNo As Long, ByRef str�������� As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:���ָ����������Ϣ
    '------------------------------------------------------------------------------------------------------------------
    
    Dim strSQL As String
    Dim blnNotice As Boolean
    Dim blnCheck As Boolean
    Dim blnHaveResult As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim str�������� As String
    Dim lngPos As Long
    Dim lngPosField As Long
    Dim strTmp As String
    Dim strTmpField As String
    Dim strField As String
    Dim strFieldType As String
    Dim strResult As String
    Dim lngLoop As Long

    
    strSQL = "SELECT A.���, A.ϵͳ, A.��������, A.��������, A.���ѱ���, A.��������, A.���Ѵ���, A.����˳��, A.�������, A.��������, A.��ʼʱ��, A.��ֹʱ��,B.���ʱ��,B.����ʱ��," & _
                    "DECODE(B.���ʱ��,NULL,NULL,DECODE(�������,NULL,NULL,SYSDATE - (B.���ʱ��+�������/(24*60)))) AS �������," & _
                    "DECODE(B.����ʱ��,NULL,NULL,DECODE(��������,NULL,NULL,SYSDATE - (B.����ʱ��+��������/(24*60)))) AS �������� " & _
                    "FROM zlNotices A,(SELECT �������, �û���, ���ʱ��, �����, ���ѱ�־, ����ʱ��, �Ѷ���־, �������� FROM zlNoticeRec WHERE �û���=[1])B " & _
                    "WHERE A.���=B.�������(+) AND A.���=" & lngNo

    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strUser)
    If rs.BOF Then Exit Function
            
    blnCheck = False
        
    If IsNull(rs("���ʱ��").Value) Or IsNull(rs("�������").Value) Then
        blnCheck = True
    Else
        If IsNull(rs("�������").Value) = False Then
            If rs("�������").Value >= 0 Then blnCheck = True
        End If
    End If
    
    If blnCheck Then
    
        blnHaveResult = False
        str�������� = ""
        
        If IsNull(rs("��������").Value) = False Then
            
            str�������� = UCase(rs("��������").Value)
            
            '�滻���������еĹ̶�����[USER]
            
            str�������� = ReplaceAll(str��������, "[USER]", "[1]")
                
            '��Ϊ��ʱ����ִ��SQL���ɹ�,û��Ȩ��,�򲻶Ա���Ϣ���д���
            Err = 0
            On Error GoTo errHand
            
            Set rsTmp = zlDatabase.OpenSQLRecord(str��������, Me.Caption, strUser)
            
            str�������� = ""
            
            If Err = 0 Then
                If rsTmp.BOF = False Then
                    blnHaveResult = True
                    
                    'str�������� = NVL(rsTmp("���").Value)
                                        
                Else
                    '��������������,ɾ����¼
                    strSQL = "Zl_Zlnoticerec_Edit(2," & lngNo & ",'" & strUser & "',Null,Null,Null,Null,Null)"
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                    
                    Exit Function
                End If
            Else
                Err = 0
                
                Exit Function
                
            End If
            
            
            '������������,���滻���õ��ֶ�ֵ
            
            str�������� = NVL(rs("��������").Value)
            
            'strTmp��ʽ:��'[����];varchar2|[�Ա�];date'
            strTmp = NVL(rs("����˳��").Value) & "|"
                           
            Do While strTmp <> "|" And strTmp <> ""
                
                lngPos = InStr(strTmp, "|")
                strTmpField = Mid(strTmp, 1, lngPos - 1)
                lngPosField = InStr(strTmpField, ";")
                
                If lngPosField > 0 Then
                
                    strField = Mid(strTmpField, 1, lngPosField - 1)
                    strFieldType = Trim(UCase(Mid(strTmpField, lngPosField + 1)))
                    
                    strTmp = Trim(Mid(strTmp, lngPos + 1))
                    
                    lngPos = InStr(str��������, strField)
                    
                    If lngPos > 0 Then
                        
                        strResult = Trim(Mid(strField, 2))
                        strResult = Mid(strResult, 1, Len(strResult) - 1)
                        
                        On Error Resume Next
                        Err = 0
                        If strFieldType = "NUMBER" Then
                        
                            strResult = NVL(rsTmp(strResult).Value)
                            
'                            strResult = "to_char(" & strResult & ")"
                        ElseIf strFieldType = "DATE" Then
'                            strResult = "to_char(" & strResult & ",'yyyy-mm-dd')"
                            strResult = Format(NVL(rsTmp(strResult).Value), "yyyy-MM-dd")
                        Else
                            strResult = NVL(rsTmp(strResult).Value)
                        End If
                        
                        'str�������� = Trim(Mid(str��������, 1, lngPos - 1) & "'||" & strResult & "||'" & Mid(str��������, lngPos + Len(strField)))
                        
                        If Err = 0 Then
                            str�������� = Trim(Mid(str��������, 1, lngPos - 1) & strResult & Mid(str��������, lngPos + Len(strField)))
                        End If
                        
                        On Error GoTo errHand
                        
                    End If
                    
                End If
                
            Loop
            
'            lngPos = InStr(str��������, " FROM ")
'
'            If lngPos > 0 Then
'                strSQL = Trim("SELECT '" & str�������� & "' AS ��� " & Mid(str��������, lngPos))
'
'                '��Ϊ��ʱ����ִ��SQL���ɹ�,û��Ȩ��,�򲻶Ա���Ϣ���д���
'                Err = 0
'                On Error GoTo errHand
'
'                If rsTmp.State = adStateOpen Then rsTmp.Close
'                rsTmp.Open strSQL, gcnOracle
'
'
'                str�������� = ""
'
'                If Err = 0 Then
'                    If rsTmp.BOF = False Then
'                        blnHaveResult = True
'                        str�������� = NVL(rsTmp("���").Value)
'                    Else
'                        '��������������,ɾ����¼
'                        strSQL = "DELETE FROM zlNoticeRec WHERE �������=" & lngNo & " AND upper(�û���)='" & UCase(strUser) & "'"
'                        gcnOracle.Execute strSQL
'
'                        Exit Function
'                    End If
'                Else
'                    Err = 0
'
'                    Exit Function
'
'                End If
'            End If
        Else
            blnHaveResult = True
            str�������� = NVL(rs("��������").Value)
        End If
                            
        strSQL = "Zl_Zlnoticerec_Edit(2," & lngNo & ",'" & strUser & "',Null,Null,Null,Null,Null)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        strSQL = "Zl_Zlnoticerec_Edit(0," & lngNo & ",'" & UCase(strUser) & "',SYSDATE," & IIf(blnHaveResult, 1, 0) & ",0,0,'" & str�������� & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    
    '���Ѽ��
    blnNotice = False
        
    If IsNull(rs("�������").Value) Then
        blnNotice = True
    Else
        If IsNull(rs("����ʱ��").Value) Then
            blnNotice = True
        Else
            If IsNull(rs("��������").Value) = False Then
                If rs("��������").Value >= 0 Then blnNotice = True
            End If
        End If
    End If
    
    If blnNotice Then
        strSQL = "Zl_Zlnoticerec_Edit(1," & lngNo & ",'" & strUser & "',Null,Null,1,Null,Null)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
            
    CheckNotice = blnNotice
    
    Exit Function
    
errHand:
    
End Function

Private Function GetClientIP(ByVal strWorkStation As String) As String
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    GetClientIP = strWorkStation
    
    strSQL = "Select IP From zlClients Where ����վ=[1] And IP Is Not Null"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(strWorkStation))
    If rs.BOF = False Then
        GetClientIP = rs("IP").Value
    End If
    
End Function
Private Function SendMessage(ByVal str����վ As String, _
                            ByVal lng�˿ں� As Long, _
                            ByVal str�û��� As String, _
                            ByVal str��Ϣ�� As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:����������Ϣ
    '------------------------------------------------------------------------------------------------------------------
    
    Dim strTmp As String
    Dim varTmp As Variant
    Dim varTmp2 As Variant
    Dim lngCount As Long
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    If Trim(str��Ϣ��) = "" Then Exit Function
'
'    'Ҫ���͵Ŀͻ��˻�������(IP)�Ͷ˿ں�
    winSock.RemoteHost = str����վ
    winSock.RemotePort = lng�˿ں�
    
    '������Ϣ
    winSock.SendData str��Ϣ��
    
    SendMessage = True
    
    Exit Function
    
errHand:
    
End Function


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

