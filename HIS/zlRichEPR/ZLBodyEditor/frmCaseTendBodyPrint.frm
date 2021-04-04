VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCaseTendBodyPrint 
   Caption         =   "���±�Ԥ��"
   ClientHeight    =   6090
   ClientLeft      =   165
   ClientTop       =   150
   ClientWidth     =   9030
   Icon            =   "frmCaseTendBodyPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   9030
   StartUpPosition =   1  '����������
   Begin VB.VScrollBar scrVsc 
      DragIcon        =   "frmCaseTendBodyPrint.frx":5C12
      Height          =   4755
      LargeChange     =   20
      Left            =   8775
      Max             =   100
      MouseIcon       =   "frmCaseTendBodyPrint.frx":5F1C
      SmallChange     =   10
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   735
      Width           =   250
   End
   Begin VB.HScrollBar scrHsc 
      DragIcon        =   "frmCaseTendBodyPrint.frx":606E
      Height          =   250
      LargeChange     =   20
      Left            =   0
      Max             =   100
      MouseIcon       =   "frmCaseTendBodyPrint.frx":6378
      SmallChange     =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5475
      Width           =   8760
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4755
      Left            =   15
      ScaleHeight     =   4755
      ScaleWidth      =   8760
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   8760
      Begin VB.PictureBox picShow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Left            =   270
         MouseIcon       =   "frmCaseTendBodyPrint.frx":64CA
         MousePointer    =   99  'Custom
         ScaleHeight     =   3390
         ScaleWidth      =   6990
         TabIndex        =   6
         Top             =   195
         Visible         =   0   'False
         Width           =   6990
      End
      Begin VB.PictureBox picPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Index           =   0
         Left            =   7815
         MouseIcon       =   "frmCaseTendBodyPrint.frx":661C
         MousePointer    =   99  'Custom
         ScaleHeight     =   3390
         ScaleWidth      =   6990
         TabIndex        =   2
         Top             =   300
         Width           =   6990
      End
      Begin VB.PictureBox picShadow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Left            =   330
         ScaleHeight     =   3390
         ScaleWidth      =   6990
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   255
         Width           =   6990
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5730
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendBodyPrint.frx":676E
            Text            =   "��������"
            TextSave        =   "��������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13018
            Object.ToolTipText     =   "��ӡ����Ϣ"
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
      DesignerControls=   "frmCaseTendBodyPrint.frx":7002
   End
End
Attribute VB_Name = "frmCaseTendBodyPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintCurPage As Integer
Private mintPage As Integer
Private mlngWidth As Long, mlngHeight As Long
Private mlngPreX As Long, mlngPreY As Long
Private Const Shadow_W = 45 '��Ӱ���
Private mintPrintRange As Integer
Private mlngBeginY As Long
Private mlngLeft As Long
Private mstrPrinter As String
Private mlng����id As Long
Private mlngPatiId As Long            '����ID
Private mlngPageId As Long            '��ҳID
Private mlngSectID As String          '����ID
Private msngScale As Single
Private mintBeginPage As Integer      '�ӵڼ�ҳ�ſ�ʼ��ӡ
Private mintEndPage As Integer        '���ڼ�ҳ�Ž�����ӡ
Private mintPageNo As Integer         '��ʼ��ӡҳ��
Private mstrArrFromTo() As String
Private mlngStartIndex As Long
Private strSQL As String
Private mlngCaseRecordID As Long
Private rsTemp As New ADODB.Recordset
Private mobjCombo As CommandBarComboBox
Private mobjPage As CommandBarComboBox

Public Event AfterPrint()

Public Function Preview(ByVal intPrintRange As Integer, ByVal lngBeginY As Long, ByVal lngLeft As Long, _
    ByVal objParent As Object, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngSectID As Long, _
    intBeginPage As Integer, intEndPage As Integer, intPageNo As Integer, strArrFromTo() As String, Optional ByVal lngStartIndex As Long = -1, Optional ByVal lng����id As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ���ָ���Ĳ���(��)���д�ӡԤ��
    '������lngBeginY=��ʼ���λ��(mm)
    '      objParent=�����ô���
    '******************************************************************************************************************
    On Error GoTo ErrHandle
    
    Dim i As Long
    
    msngScale = 1
    mintPrintRange = intPrintRange

    mlngPatiId = lngPatiID
    mlngPageId = lngPageId

    mlngSectID = CStr(lngSectID)
        
    mlngBeginY = lngBeginY
    mlng����id = lng����id
    mintPageNo = intPageNo
    mintBeginPage = intBeginPage
    mintEndPage = intEndPage
    
    mlngStartIndex = lngStartIndex
    
    ReDim Preserve mstrArrFromTo(UBound(strArrFromTo))
    
    For i = 0 To UBound(strArrFromTo)
        mstrArrFromTo(i) = strArrFromTo(i)
    Next
    
    Call InitMenuBar
    
'    Set mobjParent = objParent
    
    If picPage.UBound >= 0 Then Call ShowPage(mintCurPage, msngScale)
    
    Me.Show vbModal, objParent
    
    Preview = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitMenuBar() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ���˵���������
    '------------------------------------------------------------------------------------------------------------------
    Dim cbrMenuBar As Object
    Dim obj As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim intLoop As Integer
    
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '
    '------------------------------------------------------------------------------------------------------------------
    cbsThis.ActiveMenuBar.Title = "�˵���"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
    '
    '------------------------------------------------------------------------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)")
    With cbrMenuBar.CommandBar.Controls
        .Add xtpControlButton, conMenu_File_Print, "��ӡ(&P)..."
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
        cbrControl.BeginGroup = True
    End With
    
    '
    '------------------------------------------------------------------------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "��ͼ(&V)")
    With cbrMenuBar.CommandBar.Controls
            
        Set objPopup = .Add(xtpControlPopup, 0, "���ű���(&C)")
        objPopup.BeginGroup = True

        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, 1, "ʵ�ʴ�С(&A)"): cbrControl.Parameter = "1.00"
        cbrControl.BeginGroup = True
                
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, 2, "75%"): cbrControl.Parameter = "0.75"
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, 2, "50%"): cbrControl.Parameter = "0.50"
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, 2, "30%"): cbrControl.Parameter = "0.30"
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, 2, "25%"): cbrControl.Parameter = "0.25"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Navigatebeginning, "��һҳ(&F)")
        cbrControl.BeginGroup = True
        .Add xtpControlButton, conMenu_View_Navigateleft, "ǰһҳ(&P)"
        .Add xtpControlButton, conMenu_View_Navigateright, "��һҳ(&N)"
        .Add xtpControlButton, conMenu_View_Navigateend, "���һҳ(&L)"
        
    End With
    
    '
    '------------------------------------------------------------------------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.Id = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        objPopup.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        objPopup.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
        objPopup.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)..."): cbrControl.BeginGroup = True
    End With
    
    '����������
    '------------------------------------------------------------------------------------------------------------------
    Set cbrToolBar = cbsThis.Add("��ӡԤ��", xtpBarTop)
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ"): cbrControl.Style = xtpButtonIconAndCaption
    
        Set cbrControl = .Add(xtpControlButton, 1, "ʵ�ʴ�С")
        cbrControl.BeginGroup = True
        cbrControl.Parameter = "1.00"
        
        Set mobjCombo = .Add(xtpControlComboBox, 3, "")
        mobjCombo.BeginGroup = True
        mobjCombo.AddItem "100%", 1
        mobjCombo.AddItem "75%", 2
        mobjCombo.AddItem "50%", 3
        mobjCombo.AddItem "30%", 4
        mobjCombo.AddItem "25%", 5
        mobjCombo.ListIndex = 1
        mobjCombo.Width = 80
        mobjCombo.DropDownWidth = 80
        mobjCombo.DropDownListStyle = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Navigatebeginning, "��һҳ"): cbrControl.BeginGroup = True: cbrControl.Style = xtpButtonIcon
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Navigateleft, "ǰһҳ"): cbrControl.Style = xtpButtonIcon
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Navigateright, "��һҳ"): cbrControl.Style = xtpButtonIcon
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Navigateend, "���һҳ"): cbrControl.Style = xtpButtonIcon
                
        Set mobjPage = .Add(xtpControlComboBox, 6, "")
        mobjPage.BeginGroup = True
        
        For intLoop = 0 To picPage.UBound
            mobjPage.AddItem "�� " & intLoop + 1 & " ҳ", intLoop + 1
        Next
        mobjPage.ListIndex = 1
        mobjPage.Width = 80
        mobjPage.DropDownWidth = 80
        mobjPage.DropDownListStyle = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&E)"): cbrControl.Style = xtpButtonIconAndCaption
        cbrControl.BeginGroup = True
    End With
    
    
    '�����
    '------------------------------------------------------------------------------------------------------------------
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function


Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    
    Select Case Control.Id
    Case conMenu_File_Print
        Call PrintPage
        
    Case 1
        
        msngScale = 1
        mobjCombo.Text = "100%"
        
        Call ShowPage(mintCurPage, msngScale)
        
    Case 2
        
        msngScale = Val(Control.Parameter)
        mobjCombo.Text = CStr(msngScale * 100) & "%"
        
        Call ShowPage(mintCurPage, msngScale)
                
    Case 3
        
        If Val(Control.Text) > 0 Then
            If Val(Control.Text) <= 200 Then
                msngScale = Val(Control.Text) / 100
                Control.Text = CStr(Val(Control.Text)) & "%"
            Else
                If Control.ListIndex = 0 Then
                    Control.ListIndex = 1
                End If
                
                msngScale = Val(Control.List(Control.ListIndex)) / 100
                Control.Text = Control.List(Control.ListIndex)
            End If
        Else
            If Control.ListIndex = 0 Then
                Control.ListIndex = 5
            End If
                
            msngScale = Val(Control.List(Control.ListIndex)) / 100
            Control.Text = Control.List(Control.ListIndex)
        End If
        
        Call ShowPage(mintCurPage, msngScale)
        
        DoEvents
        Control.SetFocus
        
    Case conMenu_View_Navigatebeginning
        
        mintCurPage = 0
        mobjPage.ListIndex = mintCurPage + 1
        Call ShowPage(mintCurPage, msngScale)
        
    Case conMenu_View_Navigateleft
        
        If mintCurPage - 1 >= 0 Then
            mintCurPage = mintCurPage - 1
            mobjPage.ListIndex = mintCurPage + 1
            Call ShowPage(mintCurPage, msngScale)
        End If
        
    Case conMenu_View_Navigateright
        
        If mintCurPage + 1 <= picPage.UBound Then
            mintCurPage = mintCurPage + 1
            mobjPage.ListIndex = mintCurPage + 1
            Call ShowPage(mintCurPage, msngScale)
        End If
        
    Case conMenu_View_Navigateend
        
        If mintCurPage <> picPage.UBound Then
            mintCurPage = picPage.UBound
            mobjPage.ListIndex = mintCurPage + 1
            Call ShowPage(mintCurPage, msngScale)
        End If
    Case 6
        
        If Val(Control.Text) > 0 Then
            If Val(Control.Text) <= mobjPage.ListCount Then
                Control.ListIndex = Val(Control.Text)
            End If
            mintCurPage = Val(Control.ListIndex - 1)
        Else
            mintCurPage = Val(Control.ListIndex - 1)
        End If
        
        Control.Text = Control.List(Control.ListIndex)
        Call ShowPage(mintCurPage, msngScale)
        
        DoEvents
        Control.SetFocus
        
    Case conMenu_View_ToolBar_Button
    
        cbsThis(2).Visible = Not cbsThis(2).Visible
        cbsThis.RecalcLayout
        
    Case conMenu_View_StatusBar
    
        sta.Visible = Not sta.Visible
        cbsThis.RecalcLayout
            
    Case conMenu_Help_Help
    
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    
    Case conMenu_Help_About
        
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        
    Case conMenu_Help_Web_Home
        
        Call zlHomePage(Me.hWnd)
        
    Case conMenu_Help_Web_Forum '������̳
    
        Call zlWebForum(Me.hWnd)
        
    Case conMenu_Help_Web_Mail
        
        Call zlMailTo(Me.hWnd)
    
    Case conMenu_File_Exit
        Unload Me
        Exit Sub
            
    End Select
    
End Sub

'Private Sub cboPage_Click()
'
'    Call ShowPage(cboPage.ListIndex, msngScale)
'
'End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If sta.Visible Then Bottom = sta.Height
End Sub

Private Sub cbsThis_Resize()
    
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '�ͻ�����Ĵ�С

    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    picBack.Move lngLeft, lngTop, lngRight - lngLeft - scrVsc.Width, lngBottom - lngTop - scrHsc.Height
    scrVsc.Move picBack.Left + picBack.Width, lngTop, scrVsc.Width, picBack.Height
    scrHsc.Move lngLeft, picBack.Top + picBack.Height, picBack.Width, scrHsc.Height

    picShadow.Width = picShow.Width
    picShadow.Height = picShow.Height
    
    '����Ԥ��ҳ

    If picBack.ScaleWidth >= picShow.Width + Shadow_W * 4 Then
        picShow.Left = (picBack.ScaleWidth - (picShow.Width + Shadow_W * 4)) / 2 + Shadow_W * 2
        picShadow.Left = picShow.Left + Shadow_W
        scrHsc.Enabled = False
    Else
        scrHsc.Max = (picShow.Width + Shadow_W * 4 - picBack.ScaleWidth) / 15
        If scrHsc.Max / 3 < scrHsc.SmallChange Then
            scrHsc.LargeChange = scrHsc.SmallChange
        Else
            scrHsc.LargeChange = scrHsc.Max / 3
        End If
        scrHsc.Value = 0
        scrHsc.Enabled = True
        scrhsc_Change
    End If

    If picBack.ScaleHeight >= picShow.Height + Shadow_W * 4 Then
        picShow.Top = (picBack.ScaleHeight - (picShow.Height + Shadow_W * 4)) / 2 + Shadow_W
        picShadow.Top = picShow.Top + Shadow_W
        scrVsc.Enabled = False
    Else
        scrVsc.Max = (picShow.Height + Shadow_W * 4 - picBack.ScaleHeight) / 15
        If scrVsc.Max / 3 < scrVsc.SmallChange Then
            scrVsc.LargeChange = scrVsc.SmallChange
        Else
            scrVsc.LargeChange = scrVsc.Max / 3
        End If
        scrVsc.Value = 0
        scrVsc.Enabled = True
        scrVsc_Change
    End If
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
    Case 1, 2
        Control.Checked = (Val(Control.Parameter) = Val(msngScale))
        
    Case conMenu_View_Navigatebeginning
        
        Control.Enabled = picPage.UBound > 0 And mintCurPage <> 0
        
    Case conMenu_View_Navigateleft
        
        Control.Enabled = picPage.UBound > 0 And mintCurPage > 0
        
    Case conMenu_View_Navigateright
        
        Control.Enabled = picPage.UBound > 0 And mintCurPage < picPage.UBound
        
    Case conMenu_View_Navigateend
        
        Control.Enabled = picPage.UBound > 0 And mintCurPage <> picPage.UBound
        
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            If scrVsc.Enabled And scrVsc.Value > scrVsc.Min Then
                If Shift = 2 Then
                    scrVsc.Value = IIf(scrVsc.Value - scrVsc.LargeChange < scrVsc.Min, scrVsc.Min, scrVsc.Value - scrVsc.LargeChange)
                Else
                    scrVsc.Value = IIf(scrVsc.Value - scrVsc.SmallChange < scrVsc.Min, scrVsc.Min, scrVsc.Value - scrVsc.SmallChange)
                End If
            End If
        Case vbKeyDown
            If scrVsc.Enabled And scrVsc.Value < scrVsc.Max Then
                If Shift = 2 Then
                    scrVsc.Value = IIf(scrVsc.Value + scrVsc.LargeChange > scrVsc.Max, scrVsc.Max, scrVsc.Value + scrVsc.LargeChange)
                Else
                    scrVsc.Value = IIf(scrVsc.Value + scrVsc.SmallChange > scrVsc.Max, scrVsc.Max, scrVsc.Value + scrVsc.SmallChange)
                End If
            End If
        Case vbKeyLeft
            If scrHsc.Enabled And scrHsc.Value > scrHsc.Min Then
                If Shift = 2 Then
                    scrHsc.Value = IIf(scrHsc.Value - scrHsc.LargeChange < scrHsc.Min, scrHsc.Min, scrHsc.Value - scrHsc.LargeChange)
                Else
                    scrHsc.Value = IIf(scrHsc.Value - scrHsc.SmallChange < scrHsc.Min, scrHsc.Min, scrHsc.Value - scrHsc.SmallChange)
                End If
            End If
        Case vbKeyRight
            If scrHsc.Enabled And scrHsc.Value < scrHsc.Max Then
                If Shift = 2 Then
                    scrHsc.Value = IIf(scrHsc.Value + scrHsc.LargeChange > scrHsc.Max, scrHsc.Max, scrHsc.Value + scrHsc.LargeChange)
                Else
                    scrHsc.Value = IIf(scrHsc.Value + scrHsc.SmallChange > scrHsc.Max, scrHsc.Max, scrHsc.Value + scrHsc.SmallChange)
                End If
            End If
        
        Case vbKeyHome
            mintCurPage = 0
            mobjPage.ListIndex = mintCurPage + 1
            Call ShowPage(mintCurPage, msngScale)
        Case vbKeyEnd
            If mintCurPage <> picPage.UBound Then
                mintCurPage = picPage.UBound
                mobjPage.ListIndex = mintCurPage + 1
                Call ShowPage(mintCurPage, msngScale)
            End If
        Case vbKeyPageUp
            If mintCurPage - 1 >= 0 Then
                mintCurPage = mintCurPage - 1
                mobjPage.ListIndex = mintCurPage + 1
                Call ShowPage(mintCurPage, msngScale)
            End If
        Case vbKeyPageDown
            If mintCurPage + 1 <= picPage.UBound Then
                mintCurPage = mintCurPage + 1
                mobjPage.ListIndex = mintCurPage + 1
                Call ShowPage(mintCurPage, msngScale)
            End If
    End Select
End Sub

Private Sub PrintPage()
'���ܣ���ӡ���±�
    Dim i As Long
    Dim intCol As Integer
    
    If Not ExistsPrinter Then MsgBox "ϵͳ��û�п��õĴ�ӡ����", vbInformation: Exit Sub
    If MsgBox("׼����ӡ���µ�����ӡ���Ƿ��Ѿ�׼��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
    If Printer.DeviceName <> mstrPrinter Then
        For i = 0 To Printers.Count - 1
            If Printers(i).DeviceName = mstrPrinter Then Set Printer = Printers(i): Exit For
        Next
    End If
        
    On Error Resume Next
    'ֽ��
    If mintPage = 256 Then
        Printer.PaperSize = 256
        Printer.Width = mlngWidth
        Printer.Height = mlngHeight
    Else
        Printer.PaperSize = mintPage
        mlngWidth = Printer.Width
        mlngHeight = Printer.Height
    End If
    
    On Error GoTo ErrPrintState

       'WinNT�Զ���ֽ�Ŵ���
    If IsWindowsNT And mintPage = 256 Then
        If AddCustomPaper(Me.hWnd, mlngWidth / 56.7, mlngHeight / 56.7) = FORM_NOT_SELECTED Then Exit Sub
    End If


    '���ô�ӡ
    Select Case mintPrintRange
    Case 0
        
        Printer.PaintPicture picPage(picPage.UBound).Image, 0, 0
        Printer.EndDoc
        
    Case 1, 2
        
        For intCol = mlngStartIndex To UBound(mstrArrFromTo)
        
            Printer.PaintPicture picPage(intCol - mlngStartIndex).Image, 0, 0
            If intCol = UBound(mstrArrFromTo) Then
                Printer.EndDoc
            Else
                Printer.NewPage
            End If
        Next
                
    End Select
    
    On Error Resume Next
    If mintPage = 256 Then
        Printer.PaperSize = 256
        Printer.Width = mlngWidth
        Printer.Height = mlngHeight
    Else
        Printer.PaperSize = mintPage
        mlngWidth = Printer.Width
        mlngHeight = Printer.Height
    End If
'    WinNT�Զ���ֽ�Ŵ���
    If IsWindowsNT And mintPage = 256 Then DelCustomPaper
        
    Call ShowPrinterInfo
    
    RaiseEvent AfterPrint
    Exit Sub
ErrPrintState:
    
    MsgBox "��ӡ����ʼ��ʧ�ܣ�", vbExclamation, gstrSysName
End Sub

Private Sub picback_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As Object
    
    If Button <> 2 Then Exit Sub
    
    If cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = cbsThis.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        If cbrControl.Id > 0 Then
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.Id, cbrControl.Caption)
            cbrPopupItem.IconId = cbrControl.IconId
            cbrPopupItem.Parameter = cbrControl.Parameter
            cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        End If
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub picback_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If scrVsc.Enabled Then
            If (y - mlngPreY) / 15 > 0 Then
                scrVsc.Value = IIf(scrVsc.Value - (y - mlngPreY) / 15 < scrVsc.Min, scrVsc.Min, scrVsc.Value - (y - mlngPreY) / 15)
            Else
                scrVsc.Value = IIf(scrVsc.Value - (y - mlngPreY) / 15 > scrVsc.Max, scrVsc.Max, scrVsc.Value - (y - mlngPreY) / 15)
            End If
        End If
        If scrHsc.Enabled Then
            If (x - mlngPreX) / 15 > 0 Then
                scrHsc.Value = IIf(scrHsc.Value - (x - mlngPreX) / 15 < scrHsc.Min, scrHsc.Min, scrHsc.Value - (x - mlngPreX) / 15)
            Else
                scrHsc.Value = IIf(scrHsc.Value - (x - mlngPreX) / 15 > scrHsc.Max, scrHsc.Max, scrHsc.Value - (x - mlngPreX) / 15)
            End If
        End If
    End If
End Sub

Private Sub picShow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As Object
    
    If Button = 1 Then Set picShow.MouseIcon = scrHsc.MouseIcon
    If Button <> 2 Then Exit Sub
    
    If cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls(2)
        
    Set cbrPopupBar = cbsThis.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        If cbrControl.Id > 0 Then
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.Id, cbrControl.Caption)
            cbrPopupItem.IconId = cbrControl.IconId
            cbrPopupItem.Parameter = cbrControl.Parameter
            cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        End If
    Next
    cbrPopupBar.ShowPopup
    
End Sub

Private Sub picPage_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If scrVsc.Enabled Then
            If (y - mlngPreY) / 15 > 0 Then
                scrVsc.Value = IIf(scrVsc.Value - (y - mlngPreY) / 15 < scrVsc.Min, scrVsc.Min, scrVsc.Value - (y - mlngPreY) / 15)
            Else
                scrVsc.Value = IIf(scrVsc.Value - (y - mlngPreY) / 15 > scrVsc.Max, scrVsc.Max, scrVsc.Value - (y - mlngPreY) / 15)
            End If
        End If
        If scrHsc.Enabled Then
            If (x - mlngPreX) / 15 > 0 Then
                scrHsc.Value = IIf(scrHsc.Value - (x - mlngPreX) / 15 < scrHsc.Min, scrHsc.Min, scrHsc.Value - (x - mlngPreX) / 15)
            Else
                scrHsc.Value = IIf(scrHsc.Value - (x - mlngPreX) / 15 > scrHsc.Max, scrHsc.Max, scrHsc.Value - (x - mlngPreX) / 15)
            End If
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub picPage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Set picPage(Index).MouseIcon = scrVsc.MouseIcon
End Sub

Private Sub scrVsc_Change()
    picShow.Top = -scrVsc.Value * 15# + Shadow_W * 2
    picShadow.Top = picShow.Top + Shadow_W
    Me.Refresh
End Sub

Private Sub scrVsc_Scroll()
    picShow.Top = -scrVsc.Value * 15# + Shadow_W * 2
    picShadow.Top = picShow.Top + Shadow_W
    Me.Refresh
End Sub

Private Sub scrhsc_Change()
    picShow.Left = -scrHsc.Value * 15# + Shadow_W * 2
    picShadow.Left = picShow.Left + Shadow_W
    Me.Refresh
End Sub

Private Sub scrhsc_Scroll()
    picShow.Left = -scrHsc.Value * 15# + Shadow_W * 2
    picShadow.Left = picShow.Left + Shadow_W
    Me.Refresh
End Sub

Private Sub Form_Load()
    mintCurPage = 0
    RestoreWinState Me, App.ProductName
    
    mlngLeft = Val(zlDatabase.GetPara("���µ���߾�", glngSys, 1255, OFFSET_LEFT)) * 56.7
    mlngWidth = Val(zlDatabase.GetPara("���µ�����", glngSys, 1255, Printer.Width))
    mlngHeight = Val(zlDatabase.GetPara("���µ��߶�", glngSys, 1255, Printer.Height))
    mintPage = Val(zlDatabase.GetPara("���µ�ֽ��", glngSys, 1255, Printer.PaperSize))
    
    Call ShowPrinterInfo
    
'    'ȱʡ��ԭʼ��С�ķ�ʽ��ʾ����
'    Call mnuView_ScaleValue_Click(0)
End Sub

Private Sub ScalePicture(ByVal sngScale As Single)
    Dim obj As StdPicture
    Dim intCol As Integer
    Dim x As Single
    Dim y As Single
    
    On Error GoTo errHand
    
    picShadow.Width = picShadow.Width * sngScale
    picShadow.Height = picShadow.Height * sngScale
    
    For intCol = mlngStartIndex To UBound(mstrArrFromTo)
        
        Set obj = picPage(intCol - mlngStartIndex).Image
        
        picPage(intCol - mlngStartIndex).Width = picPage(intCol - mlngStartIndex).Width * sngScale
        picPage(intCol - mlngStartIndex).Height = picPage(intCol - mlngStartIndex).Height * sngScale
        
        x = picPage(intCol - mlngStartIndex).Width
        y = picPage(intCol - mlngStartIndex).Height
        
        picPage(intCol - mlngStartIndex).Cls
        picPage(intCol - mlngStartIndex).PaintPicture obj, 0, 0, x, y
        
    Next
errHand:
    
End Sub

Private Function ShowPage(ByVal PageNum As Long, ByVal sngScale As Single) As Boolean
    
'    If PageNum = 0 Then Exit Function
    
    On Error GoTo errHand
    
    LockWindowUpdate picShow.hWnd
            
    picPage(0).Visible = False
    
    If picShow.Visible = False Then picShow.Visible = True
    
    picShow.Width = picPage(PageNum).Width * sngScale
    picShow.Height = picPage(PageNum).Height * sngScale
    picShow.Cls
    picBack.Cls
    
    '���ð�ɫ������Ч����ã�
    SetStretchBltMode picShow.hDC, HALFTONE
    StretchBlt picShow.hDC, 0, 0, picShow.Width, picShow.Height, picPage(PageNum).hDC, 0, 0, picPage(PageNum).Width, picPage(PageNum).Height, SRCCOPY
    
    Call cbsThis_Resize
    
'    If PageNum = mlngStartPage And mlngBlankHeight > 100 Then Call DrawAlphaRect(mlngBlankHeight * ZoomFactor)
'    Call Form_Resize
    
errHand:
    LockWindowUpdate 0
    UpdateWindow picShow.hWnd
    
End Function

Private Sub ShowPrinterInfo()
    sta.Panels(2).Text = "��ӡ��:" & Printer.DeviceName & "/ֽ��:" & _
        GetPaperName(Printer.PaperSize) & "/�ߴ�:" & _
        CLng(Printer.Width / 56.7) & "��" & CLng(Printer.Height / 56.7) & "/ֽ��:" & _
        IIf(Printer.Orientation = 1, "����", "����")
End Sub





