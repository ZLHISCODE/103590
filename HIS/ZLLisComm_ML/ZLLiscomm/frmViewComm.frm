VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmViewComm 
   Caption         =   "ͨѶ���"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11340
   Icon            =   "frmViewComm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   11340
   StartUpPosition =   2  '��Ļ����
   Begin RichTextLib.RichTextBox Rtxt_������� 
      Height          =   2460
      Left            =   6495
      TabIndex        =   0
      Top             =   3750
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   4339
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmViewComm.frx":000C
   End
   Begin RichTextLib.RichTextBox Rtxt_δ֪�� 
      Height          =   2460
      Left            =   3795
      TabIndex        =   1
      Top             =   1020
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   4339
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmViewComm.frx":00A9
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   7005
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmViewComm.frx":0146
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11509
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1905
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "2009-10-15"
            Key             =   "DATE"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "16:21"
            Key             =   "TIME"
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
      Left            =   780
      Top             =   105
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpCom 
      Bindings        =   "frmViewComm.frx":09DA
      Left            =   180
      Top             =   90
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmViewComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrDev As String '��ǰ��ص��豸
Private mstrCom As String
Private mstrCLSName As String
Private mLasterTime As Date

Public Event CloseWindow()

Public Sub ShowDecode(ByVal intType As Integer, ByVal str_In As String)
    '��ʾ�յ��Ľ�������δ֪��
    'inttype: 0-��� 1-δ֪��
    If intType = 0 Then
        Rtxt_�������.Text = Rtxt_�������.Text & IIf(Rtxt_�������.Text = "", "", vbNewLine)
        Rtxt_�������.Text = Rtxt_�������.Text & "======������======" & vbNewLine
        Rtxt_�������.Text = Rtxt_�������.Text & str_In
        Rtxt_�������.SelStart = Len(Rtxt_�������.Text)
    Else
        Rtxt_δ֪��.Text = Rtxt_δ֪��.Text & str_In
        Rtxt_δ֪��.SelStart = Len(Rtxt_δ֪��.Text)
    End If
End Sub

'Public Sub ShowInOut(ByVal intType As Integer, ByVal str_In As String, Optional ByVal lngEven As Long)
'    '���ܣ���ʾ�յ��ͷ��͵���Ϣ
'    'intType��0-���� 1-���� 2-����
'    Me.stbThis.Panels(2).Text = "������" & mstrDev & " " & mstrCom
'    Dim str�� As String
''    str�� = IIf(RTxt_Hex.Text = "", "", "(+ " & DateDiff("s", mLasterTime, Now) & " ��)")
'    If intType = 0 Then
'        Rtxt_In.Text = Rtxt_In.Text & str_In
'        Rtxt_In.SelStart = Len(Rtxt_In.Text)
''        Select Case lngEven
''        Case -1 '��IP��ʽ,�������Ǹ�onComm����һ��
''        End Select
'        RTxt_Hex.Text = RTxt_Hex.Text & IIf(RTxt_Hex.Text = "", "", vbNewLine)
'        RTxt_Hex.Text = RTxt_Hex.Text & "���գ�" & Format(Now, "yyyy-MM-dd hh:MM:ss") & str��
'        RTxt_Hex.Text = RTxt_Hex.Text & vbNewLine & str_In
'    ElseIf intType = 1 Then
'        Rtxt_Out.Text = Rtxt_Out.Text & str_In
'        Rtxt_Out.SelStart = Len(Rtxt_Out.Text)
'
'        RTxt_Hex.Text = RTxt_Hex.Text & IIf(RTxt_Hex.Text = "", "", vbNewLine)
'        RTxt_Hex.Text = RTxt_Hex.Text & "���ͣ�" & Format(Now, "yyyy-MM-dd hh:MM:ss") & str��
'        RTxt_Hex.Text = RTxt_Hex.Text & vbNewLine & str_In
'    Else
'        RTxt_Hex.Text = RTxt_Hex.Text & IIf(RTxt_Hex.Text = "", "", vbNewLine)
'        RTxt_Hex.Text = RTxt_Hex.Text & "����" & Format(Now, "yyyy-MM-dd hh:MM:ss") & str��
'        RTxt_Hex.Text = RTxt_Hex.Text & vbNewLine & str_In
'    End If
'
'
'    RTxt_Hex.SelStart = Len(RTxt_Hex.Text)
'    'richtxtHEX ��ʾ���գ����͵���Ϣ
'
'    mLasterTime = Now
'End Sub

Public Function ShowMe(ByVal strDev As String, ByVal strCOM As String, ByVal strCLSName As String)
    'strDev �豸��
    'strCom ���ڼ�ͨѶ����
    'strCLSName ͨѶ������
    mstrDev = strDev
    mstrCom = ""
    If strCOM <> "" Then
        If UBound(Split(strCOM, "|")) > 1 Then
            
            mstrCom = Split(strCOM, "|")(0)
            Select Case Split(strCOM, "|")(1)
                Case "0": mstrCom = mstrCom & " û������"
                Case "1": mstrCom = mstrCom & " (XON/XOFF) ����"
                Case "2": mstrCom = mstrCom & " RTS/CTS ����"
                Case "3": mstrCom = mstrCom & " RTS/CTS �� XON/XOFF ���ֽԿ� "
            End Select
            Select Case Split(strCOM, "|")(2)
                Case "0": mstrCom = mstrCom & " �����ı�"
                Case "1": mstrCom = mstrCom & " ���ն�����"
            End Select
        ElseIf UBound(Split(strCOM, "|")) > 0 Then
            'mstrCom = Split(strCOM, "|")(1)
            Select Case Split(strCOM, "|")(0)
                Case "0": mstrCom = mstrCom & " ��Ϊ�ն� �˿ڣ�" & Mid(Split(strCOM, "|")(1), InStr(Split(strCOM, "|")(1), ":") + 1)
                Case "1": mstrCom = mstrCom & " ��Ϊ���� �˿ڣ�" & Mid(Split(strCOM, "|")(1), InStr(Split(strCOM, "|")(1), ":") + 1)
            End Select
        End If
    End If
    mstrCLSName = strCLSName
    
    Me.Show
    
End Function

Private Sub initCbsThis(cbsMain As CommandBars)
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.Icons = frmPubIcons.imgPublic.Icons
    cbsMain.Options.LargeIcons = False
    
    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)  '����
    objMenu.ID = conMenu_FilePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&T)��")  '����
'        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
'        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True '����
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "����(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_View & 2, "���մ���(&J)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_View & 3, "���ʹ���(&F)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_View & 4, "���봰��(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_View & 5, "δ֪���(&W)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False) '����
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)") '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "���(&R)"): objControl.BeginGroup = True '����
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "�������(&C)")
    End With

'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False) '����
'    objMenu.Id = conMenu_HelpPopup
'    With objMenu.CommandBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)") '����
'
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrSysName)  '����
'        With objPopup.CommandBar.Controls
'            .Add xtpControlButton, conMenu_Help_Web_Home, gstrSysName & "��ҳ(&H)", -1, False '����
'            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrSysName & "��̳(&F)", -1, False '����
'            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False '����
'        End With
'        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True '����
'    End With

    '���������⴦��
    '-----------------------------------------------------


    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
'        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "��մ���"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "���洰��")

'        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): objControl.BeginGroup = True '����
        

    End With

    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings

'        .Add FCONTROL, vbKeyP, conMenu_File_Print           '��ӡ

'        .Add 0, vbKeyF1, conMenu_Help_Help                  '����
    End With

    '����һЩ�����Ĳ���������
    '-----------------------------------------------------
'    With cbsMain.Options
'        .AddHiddenCommand conMenu_File_PrintSet         '��ӡ����
'        .AddHiddenCommand conMenu_File_Excel            '�����Excel
'    End With

    '��ȡ��������ģ��ı���(��������ģ���)
    '-----------------------------------------------------
'    Call gobjDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)

End Sub

Private Sub dkpComInit()
    Dim objPaneA As Pane, objPaneB As Pane, objPaneC As Pane, objPaneD As Pane, objPaneE As Pane
    Dim lngX As Long
    Dim lngY As Long
    
    'DockingPane ��ʼ��
    '-----------------------------------------------------
    Me.dkpCom.SetCommandBars Me.cbsThis
     
'    Set objPaneA = Me.dkpCom.CreatePane(1, 200, 135, DockTopOf)
'    objPaneA.Title = "ͨѶ��Ϣ"
'    objPaneA.Options = PaneNoCloseable Or PaneNoFloatable
'
'    Set objPaneB = Me.dkpCom.CreatePane(2, 380, 335, DockRightOf, objPaneA)
'    objPaneB.Title = "������Ϣ"
'
'    Set objPaneC = Me.dkpCom.CreatePane(3, 380, 335, DockBottomOf, objPaneB)
'    objPaneC.Title = "������Ϣ"
    
    Set objPaneD = Me.dkpCom.CreatePane(4, 100, 135, DockTopOf)
    objPaneD.Title = "������"
    
    Set objPaneE = Me.dkpCom.CreatePane(5, 100, 135, DockRightOf, objPaneD)
    objPaneE.Title = "δ֪��Ŀ"
    
    Me.dkpCom.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpCom.Options.ThemedFloatingFrames = True
    Me.dkpCom.Options.AlphaDockingContext = False
    Me.dkpCom.Options.HideClient = True
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer
    Dim objControl As CommandBarControl
    Dim strFileName As String
    Select Case Control.ID
        Case conMenu_View_ToolBar_Button '������
            For i = 2 To cbsThis.Count
                Me.cbsThis(i).Visible = Not Me.cbsThis(i).Visible
            Next
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text '��ť����
            For i = 2 To cbsThis.Count
                For Each objControl In Me.cbsThis(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size '��ͼ��
            Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
            Me.cbsThis.RecalcLayout
        Case conMenu_View_StatusBar '״̬��
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsThis.RecalcLayout
    
        Case conMenu_Edit_ClsAll
'            RTxt_Hex.Text = ""
'            Rtxt_In.Text = ""
'            Rtxt_Out.Text = ""
            Rtxt_�������.Text = ""
            Rtxt_δ֪��.Text = ""
        Case conMenu_Edit_Delete
           If TypeName(Me.ActiveControl) = "RichTextBox" Then
                Me.ActiveControl.Text = ""
           End If
        Case conMenu_Edit_Save
           If TypeName(Me.ActiveControl) = "RichTextBox" Then
                If Me.ActiveControl.Text <> "" Then
                    strFileName = Me.ActiveControl.Name
                    strFileName = Replace(strFileName, "in", "����")
                    strFileName = Replace(strFileName, "Out", "����")
                    strFileName = Replace(strFileName, "Hex", "ͨѶ")
                    strFileName = Mid(strFileName, InStr(strFileName, "_") + 1) & Format(Now, "yyMMdd_hhMMss") & ".txt"
                    Me.ActiveControl.SaveFile App.Path & "\" & strFileName, 1
                    Me.stbThis.Panels(2).Text = "�ѱ���Ϊ" & App.Path & "\" & strFileName
                End If
           End If
        Case conMenu_Edit_Seat_View & 2
            If dkpCom.Panes(2).Closed Then
                dkpCom.ShowPane 2
            Else
                dkpCom.Panes(2).Close
            End If
        Case conMenu_Edit_Seat_View & 3
            If dkpCom.Panes(3).Closed Then
                dkpCom.ShowPane 3
            Else
                dkpCom.Panes(3).Close
            End If
        Case conMenu_Edit_Seat_View & 4
            If dkpCom.Panes(4).Closed Then
                dkpCom.ShowPane 4
            Else
                dkpCom.Panes(4).Close
            End If
        Case conMenu_Edit_Seat_View & 5
            If dkpCom.Panes(5).Closed Then
                dkpCom.ShowPane 5
            Else
                dkpCom.Panes(5).Close
            End If
        Case conMenu_File_Exit        '�˳�
            Unload Me
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '������
        If cbsThis.Count >= 2 Then
            Control.Checked = Me.cbsThis(2).Visible
        End If
    Case conMenu_View_ToolBar_Text 'ͼ������
        If cbsThis.Count >= 2 Then
            Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '��ͼ��
        Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar '״̬��
        Control.Checked = Me.stbThis.Visible
    
    Case conMenu_Edit_Seat_View & 2: Control.Checked = Not dkpCom.Panes(2).Closed
    Case conMenu_Edit_Seat_View & 3: Control.Checked = Not dkpCom.Panes(3).Closed
    Case conMenu_Edit_Seat_View & 4: Control.Checked = Not dkpCom.Panes(4).Closed
    Case conMenu_Edit_Seat_View & 5: Control.Checked = Not dkpCom.Panes(5).Closed
        
    End Select
End Sub

'--------------------------------------------------------------------

Private Sub dkpCom_AttachPane(ByVal Item As XtremeDockingPane.IPane)
'    If Item.ID = 1 Then
'        Item.Handle = RTxt_Hex.hwnd
'    ElseIf Item.ID = 2 Then
'        Item.Handle = Rtxt_In.hwnd
'    ElseIf Item.ID = 3 Then
'        Item.Handle = Rtxt_Out.hwnd
'    Else
    If Item.ID = 4 Then
        Item.Handle = Rtxt_�������.hwnd
    ElseIf Item.ID = 5 Then
        Item.Handle = Rtxt_δ֪��.hwnd
    End If
End Sub

Private Sub dkpCom_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Dim lngTop As Long, lngLeft As Long, lngRight As Long, lngBottom As Long
    Me.cbsThis.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    Top = lngTop
    Bottom = Me.ScaleHeight - lngBottom
End Sub

Private Sub dkpCom_Resize()
    Me.cbsThis.RecalcLayout
End Sub

Private Sub Form_Load()
    Call initCbsThis(cbsThis)
    Call dkpComInit
    Me.stbThis.Panels(2).Text = "������" & mstrDev & " " & mstrCom
    Me.Caption = Me.Caption & " ͨѶ����:" & mstrCLSName
    
'    RTxt_Hex.Locked = True
'    Rtxt_In.Locked = True
'    Rtxt_Out.Locked = True
    Rtxt_�������.Locked = True
    Rtxt_δ֪��.Locked = True
End Sub

Private Sub Form_Resize()
    Me.dkpCom.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RaiseEvent CloseWindow
End Sub
