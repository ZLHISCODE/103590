VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmQCAddData1 
   Caption         =   "�ʿ�����¼��"
   ClientHeight    =   8580
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9900
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQCAddData2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cbo���� 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4905
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   3200
   End
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2760
      Left            =   120
      TabIndex        =   0
      Top             =   585
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   4868
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "΢���ź�"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   119341058
      CurrentDate     =   40246
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   8205
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmQCAddData2.frx":000C
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12383
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin VSFlex8Ctl.VSFlexGrid vfgQCControl 
      Height          =   2595
      Left            =   3150
      TabIndex        =   3
      Top             =   570
      Width           =   6240
      _cx             =   11007
      _cy             =   4577
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgQCdata 
      Height          =   3105
      Left            =   150
      TabIndex        =   4
      Top             =   3525
      Width           =   9390
      _cx             =   16563
      _cy             =   5477
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   270
      Top             =   90
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmQCAddData1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPriv As String
Private mrsQCData As Recordset
Private mintFmtNum As Integer

'-----------------------------------------------------------------------------
'--- �����߼�����
'-----------------------------------------------------------------------------

Private Sub cbo����_Click()
    Dim lng����id As Long, dateValue As Date
    
    If Me.cbo����.ListIndex >= 0 Then
        lng����id = Val(Me.cbo����.ItemData(Me.cbo����.ListIndex))
        dateValue = Me.mvDate.Value
        Call GetQCControlData(Me.vfgQCControl, lng����id, dateValue)
        Call vfgQCControl_RowColChange
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer
    Dim objControl As CommandBarControl
    Select Case Control.ID
    
    Case conMenu_Edit_Modify
        
        Me.vfgQCdata.Editable = flexEDKbdMouse
        Me.vfgQCdata.SelectionMode = flexSelectionFree
        Me.cbo����.Enabled = False
        Me.vfgQCControl.Enabled = False
        Me.mvDate.Enabled = False
    Case conMenu_Edit_Untread
        '
        Me.vfgQCdata.Editable = flexEDNone
        Me.vfgQCdata.SelectionMode = flexSelectionByRow
        Me.cbo����.Enabled = True
        Me.vfgQCControl.Enabled = True
        Me.mvDate.Enabled = True
        Call RefreshData
        
    
    Case conMenu_Edit_Save
        Dim lng����id As Long, lngQCID As Long, dateCurr As Date, strGetQCVal As String
        lng����id = Val(Me.cbo����.ItemData(Me.cbo����.ListIndex))
        dateCurr = Me.mvDate.Value
        lngQCID = Val("" & vfgQCControl.TextMatrix(vfgQCControl.Row, 0))
        strGetQCVal = "" & vfgQCControl.TextMatrix(vfgQCControl.Row, 8)
        Call SaveQcData(vfgQCdata, lng����id, lngQCID, dateCurr, strGetQCVal)
        
        Me.vfgQCdata.Editable = flexEDNone
        Me.vfgQCdata.SelectionMode = flexSelectionByRow
        Me.cbo����.Enabled = True
        Me.vfgQCControl.Enabled = True
        Me.mvDate.Enabled = True
        
        Call RefreshData
        
    Case conMenu_View_Refresh
        Call RefreshData
    Case conMenu_File_Exit
        Unload Me
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
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Modify
        Control.Enabled = Not (Me.vfgQCdata.Editable = flexEDKbdMouse)
    Case conMenu_Edit_Save, conMenu_Edit_Untread
        Control.Enabled = (Me.vfgQCdata.Editable = flexEDKbdMouse)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
        
    End Select
End Sub

Private Sub mvDate_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
    Call RefreshData
End Sub

Private Sub vfgQCControl_RowColChange()
    Call RefreshData
End Sub

Private Sub cbsThis_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    With Me.mvDate
        .Top = lngTop + 45
        .Left = lngLeft + 45
    End With
    With Me.vfgQCControl
        .Top = Me.mvDate.Top
        .Left = Me.mvDate.Left + Me.mvDate.Width + 45
        .Width = lngRight - .Left - 45
        .Height = Me.mvDate.Height
    End With
    With Me.vfgQCdata
        .Left = Me.mvDate.Left
        .Width = (lngRight - lngLeft) - .Left - 45
        .Top = Me.mvDate.Top + Me.mvDate.Height + 45
        
        .Height = lngBottom - .Top - Me.stbThis.Height - 45
    End With

End Sub
Public Sub ShowMe(ByVal strPrivate As String, ByVal frmMain As Form)
    mstrPriv = strPrivate
    
    Me.Show vbModal, frmMain
End Sub

Private Sub Form_Load()
    
    '��ʼʼ���ؼ���������
    '�˵�,������
    Call initCbsThis(cbsThis)
    mstrPriv = gstrPrivs
    '״̬��
    'Call InitStatusBar
    
    '��ʼ���ؼ�
    Me.mvDate.Value = Now()
    
    'װ�������������
    Call LoadInstruments(Me.cbo����)
    
End Sub

Private Function initCbsThis(cbsMain As CommandBars) As Boolean
    '��Ϊ�Ӵ��崦��˵��Ļ�׼
    '���ܣ������ڲ˵����岿��
    '˵����
    '1.���й��еĲ˵��Ͱ�ť�����У�
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
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)  '����
    objMenu.ID = conMenu_FilePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
        'Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")  '����
        'Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        'Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        'Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True '����
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "����(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&P)")
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
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True '����

    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False) '����
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)") '����
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName) '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False '����
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False '����
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True '����
    End With

    '���������⴦��
    '-----------------------------------------------------
'    ���˵��Ҳ������ѡ��
    With cbsMain.ActiveMenuBar.Controls
        Set objControl = .Add(xtpControlLabel, conMenu_View_Dept, "����")
        objControl.ID = conMenu_View_Dept
        objControl.Flags = xtpFlagRightAlign
        
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Dept + 1, "")
        objCustom.Handle = cbo����.hwnd
        objCustom.Flags = xtpFlagRightAlign
                
    End With

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        'Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ") '����
        'Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��") '����

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "����"):
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����")

        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�") '����
        
    End With

    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        '.Add FCONTROL, vbKeyP, conMenu_File_Print           '��ӡ
        .Add 0, vbKeyF1, conMenu_Help_Help                  '����
    End With

    '����һЩ�����Ĳ���������
    '-----------------------------------------------------
    With cbsMain.Options
       ' .AddHiddenCommand conMenu_File_PrintSet         '��ӡ����
       ' .AddHiddenCommand conMenu_File_Excel            '�����Excel
    End With
    
End Function

Private Sub Form_Resize()

   ' On Error Resume Next
    
    Call cbsThis_Resize

End Sub

Private Sub vfgQCdata_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strLists As String, strValue As String
    Dim lngCount As Long
    
    With Me.vfgQCdata
    
        If Col = 0 Then Exit Sub
        If Trim(.TextMatrix(Row, Col)) = "" Then Exit Sub
        
        strLists = Trim(.TextMatrix(Row, 14)) '����
        strValue = Trim(.TextMatrix(Row, Col))
        
        If strLists = "" Then

            If InStr(strValue, "E+") > 0 And Val(strValue) > 0 Then
                .TextMatrix(Row, Col) = strValue
            Else
                mintFmtNum = Val("" & .TextMatrix(Row, 17))
                If mintFmtNum > 0 Then
                    .TextMatrix(Row, Col) = Format(Val(strValue), "0." & String(mintFmtNum, "0"))
                Else
                    .TextMatrix(Row, Col) = Format(Val(strValue), "0")
                End If
            End If
            
            Exit Sub
        End If
        For lngCount = 0 To UBound(Split(strLists, ";"))
            If .TextMatrix(Row, Col) = Split(strLists, ";")(lngCount) Then Exit Sub
        Next
'        .TextMatrix(Row, Col) = ""
    End With
'    strValue = "����ĿΪ�붨����Ŀ�������ȡֵ����(" & strLists & ")Ҫ��"
'    MsgBox strValue, vbInformation, gstrSysName
End Sub

Private Sub vfgQCdata_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vfgQCdata
        If Not .TextMatrix(.FixedRows - 1, Col) Like "��?��" Then Cancel = True
    End With
End Sub

Private Sub vfgQCdata_DblClick()
    Me.vfgQCdata.Editable = flexEDKbdMouse
    Me.vfgQCdata.SelectionMode = flexSelectionFree
    Me.cbo����.Enabled = False
    Me.vfgQCControl.Enabled = False
    Me.mvDate.Enabled = False
End Sub

Private Sub vfgQCdata_KeyDown(KeyCode As Integer, Shift As Integer)
    With vfgQCdata
        If .Editable <> flexEDNone Then
            If KeyCode = vbKeyReturn Then
                KeyCode = 0
                If .TextMatrix(.FixedRows - 1, .Col) Like "��?��" Then
                    If .Row < .Rows - 1 Then
                        .Select .Row + 1, .Col
                    ElseIf .Col < .Cols - 1 Then
                        If .TextMatrix(.FixedRows - 1, .Col + 1) Like "��?��" Then .Select .FixedRows, .Col + 1
                    End If
                End If
            End If
        End If
    End With

End Sub

Private Sub vfgQCdata_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    With vfgQCdata
        If .Editable <> flexEDNone Then
            If KeyCode = vbKeyReturn Then
                If .TextMatrix(.FixedRows - 1, .Col) Like "��?��" Then
                    If .Row < .Rows - 1 Then
                        .Select .Row + 1, .Col
                    ElseIf .Col < .Cols - 1 Then
                        If .TextMatrix(.FixedRows - 1, .Col + 1) Like "��?��" Then .Select .FixedRows, .Col + 1
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub RefreshData()
    Dim lngQCID As Long, dateValue As Date, strGetQCVal As String
    dateValue = Me.mvDate.Value
    With vfgQCControl
        lngQCID = Val("" & .TextMatrix(.Row, 0))
        strGetQCVal = "" & vfgQCControl.TextMatrix(vfgQCControl.Row, 8)
       
        Call GetQCData(vfgQCdata, lngQCID, dateValue, strGetQCVal)
    End With
End Sub


'-----------------------------------------------------------------------------
'--- ���ݴ�����
'-----------------------------------------------------------------------------
Private Function SaveQcData(ByRef vsGrid As VSFlexGrid, ByVal lngDeviceID As Long, ByVal lngQCID As Long, ByVal dateWhy As Date, ByVal strGetQCVal) As Boolean
    '��������
    Dim strsql(9) As String, intRow As Integer, lng��ĿID As Long
    Dim str���(9) As String, intCol As Integer, bln�зǿս��(9) As Boolean
    Dim strSampleNO(9) As String, lng�걾ID(9) As Long, dBegin As Date, dEnd As Date
    Dim rsTemp As ADODB.Recordset, strTmp As String, rsNo As ADODB.Recordset
    Dim blnBegin As Boolean
    Dim strQCVal As String
    On Error GoTo hErr
    
    dBegin = Format(dateWhy, "yyyy-MM-dd 00:00:00")
    dEnd = Format(dateWhy, "yyyy-MM-dd 23:59:59")
    
    With vsGrid
        .Select .FixedRows - 1, 3
        For intRow = .FixedRows To .Rows - 1
            lng��ĿID = Val("" & .TextMatrix(intRow, 13))
            If lng��ĿID > 0 Then
                
                For intCol = 3 To 11
                    If strGetQCVal = "[SCO]" Then
                        strQCVal = "^^^" & Trim("" & .TextMatrix(intRow, intCol))
                    ElseIf strGetQCVal = "[OD]" Then
                        strQCVal = "^" & Trim("" & .TextMatrix(intRow, intCol)) & "^^"
                    Else
                         strQCVal = Trim("" & .TextMatrix(intRow, intCol))
                    End If
                
                    str���(intCol - 3) = str���(intCol - 3) & "|" & lng��ĿID & "^" & strQCVal
                    
                    If Trim("" & .TextMatrix(intRow, intCol)) <> "" Then
                        bln�зǿս��(intCol - 3) = True
                    End If
                Next
            End If
        Next
    End With
    'ȡÿ�εı걾��
    For intCol = LBound(strSampleNO) To UBound(strSampleNO)
        
        
        lng�걾ID(intCol) = 0
        strSampleNO(intCol) = ""
        Call GetSampleIDNO(lngDeviceID, lngQCID, dBegin, dEnd, intCol + 1, lng�걾ID(intCol), strSampleNO(intCol))
        
'        gcnOracle.BeginTrans
'        blnBegin = True
        If lng�걾ID(intCol) <= 0 Then
            '�޶�Ӧ�걾��¼��Ҫ����,����û��¼�����ݣ�ȫ�յĲ���
            If bln�зǿս��(intCol) = True Then
                lng�걾ID(intCol) = zlDatabase.GetNextId("����걾��¼")
                gstrSql = "ZL_����걾��¼_INSERT(" & lng�걾ID(intCol) & ",NULL,'" & _
                    strSampleNO(intCol) & "',NULL,NULL," & lngDeviceID & ",NULL," & _
                    "To_Date('" & Format(dateWhy, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),NULL," & _
                    "To_Date('" & Format(dateWhy, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'" & UserInfo.���� & "'," & _
                    "Null,To_Date('" & Format(dateWhy, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'" & gstrUserName & "','0',Null,0,0)"
                zlDatabase.ExecuteProcedure gstrSql, "���������ʱ��¼"
                
            End If
        
        End If
        
        If lng�걾ID(intCol) > 0 Then
            gstrSql = "ZL_������ͨ���_BATCHUPDATE(" & lng�걾ID(intCol) & "," & _
                lngDeviceID & ",Null,Null,Null,'" & Mid(str���(intCol), 2) & "')"
            zlDatabase.ExecuteProcedure gstrSql, "����������"
            
            gstrSql = "ZL_�����ʿؼ�¼_EDIT(1," & lng�걾ID(intCol) & "," & lngQCID & ",Null,Null,Null,Null,Null,Null," & intCol + 1 & ")"
            zlDatabase.ExecuteProcedure gstrSql, "����Ϊ�ʿ�Ʒ"
        End If
'        gcnOracle.CommitTrans
        blnBegin = False
        
    Next
    
    Exit Function
hErr:
    If blnBegin Then gcnOracle.RollbackTrans
    Call ErrCenter
End Function

Private Function GetSampleIDNO(ByVal lngDevId As Long, ByVal lngQC As Long, ByVal dBegin As Date, dEnd As Date, ByVal intC As Integer, ByRef lngSampleID As Long, ByRef strSampleNO As String)
    Dim strTmp As String, rsTemp As ADODB.Recordset, rsSampleNO As ADODB.Recordset
    
    On Error GoTo errH
    strTmp = "Select a.�걾id, a.�걾���,b.����, b.�걾��, b.ˮƽ" & vbNewLine & _
            "From �����ʿؼ�¼ A, �����ʿ�Ʒ B" & vbNewLine & _
            "Where �ʿ�Ʒid(+) = b.Id And b.Id = [1] And a.����ʱ��(+) between [2] and [3] And a.���Դ���(+) = [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strTmp, Me.Caption, lngQC, dBegin, dEnd, intC)
    Do Until rsTemp.EOF
        lngSampleID = Val("" & rsTemp!�걾ID)
        strSampleNO = IIf(lngSampleID <= 0, Trim("" & rsTemp!�걾��), Trim("" & rsTemp!�걾���))
        If strSampleNO = "" Or strSampleNO = "0" Then strSampleNO = rsTemp!���� & "-" & (intC - 1)
        If lngSampleID <= 0 Then
            
            Call GenNo(lngDevId, intC - 1, dBegin, dEnd, rsTemp!����, strSampleNO)
            
        End If
        rsTemp.MoveNext
    Loop
            
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub GenNo(ByVal lngDevId As Long, intC As Integer, dBegin As Date, dEnd As Date, strName As String, strSampleNO As String)
    Dim strTmp As String, rsTemp As ADODB.Recordset, rsSampleNO As ADODB.Recordset
    
    strTmp = "Select ���Դ��� from �����ʿؼ�¼ where ����ID=[1] and ����ʱ�� between [2] and [3] And �걾���=[4] "
    Set rsSampleNO = zlDatabase.OpenSQLRecord(strTmp, Me.Caption, lngDevId, dBegin, dEnd, strSampleNO)
    If Not rsSampleNO.EOF Then
        strSampleNO = strName & "-" & intC + 1
        Call GenNo(lngDevId, intC + 1, dBegin, dEnd, strName, strSampleNO)
    End If
End Sub

Private Sub LoadInstruments(ctrCbo As ComboBox, Optional intIndex As Integer)
    ' ȡ�����������ݵ�Cbo�ؼ�
    Dim strsql As String, rsTemp As ADODB.Recordset
    Dim lngMachineID As Long, lngIndex As Long
    On Error GoTo hErr
    
    lngMachineID = Val(zlDatabase.GetPara("����", glngSys, 1209, 0))
    If intIndex <> 0 Then lngIndex = intIndex
    
    If InStr(1, mstrPriv, "���п���") > 0 Then
        strsql = " Select Distinct  a.id,a.���� , a.����  From �������� a ,���ű� b,�����ʿ�Ʒ c " & _
                  "Where a.ʹ��С��ID = b.ID and a.id = c.����id"
        Set rsTemp = zlDatabase.OpenSQLRecord(strsql, gstrSysName)
        
    Else
        strsql = " Select Distinct a.id,a.���� , a.����  From ������Ա D,�������� a ,���ű� b , �����ʿ�Ʒ c " & _
                  " Where a.ʹ��С��ID = b.ID and a.ʹ��С��id=D.����id and D.��Աid = [1]  " & _
                  " and a.id = c.����Id "
        Set rsTemp = zlDatabase.OpenSQLRecord(strsql, gstrSysName, UserInfo.ID)
    End If
    
    ctrCbo.Clear
    Do Until rsTemp.EOF
        ctrCbo.AddItem "" & rsTemp!���� & " " & rsTemp!����
        ctrCbo.ItemData(ctrCbo.NewIndex) = rsTemp!ID
        If lngMachineID = rsTemp!ID Then lngIndex = ctrCbo.NewIndex
        rsTemp.MoveNext
    Loop
    
    If ctrCbo.ListCount > 0 Then
        If lngIndex >= 0 And lngIndex < ctrCbo.ListCount Then
            ctrCbo.ListIndex = lngIndex
        Else
            ctrCbo.ListIndex = 0
        End If
    End If
    Exit Sub
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetQCControlData(ByRef vsGrid As VSFlexGrid, ByVal lng����id As Long, ByVal dateWhy As Date) As Boolean
    'ȡQCControl�ؼ�������
    Dim strsql As String, rsTemp As ADODB.Recordset
    On Error GoTo hErr
    
    strsql = "Select distinct ID,�걾��,ˮƽ, ����, ����, Ũ��,  To_Char(��ʼ����, 'yyyy-MM-dd') As ��ʼ����, To_Char(��������, 'yyyy-MM-dd') As ��������,b.�ʿ�ȡֵ   " & vbNewLine & _
            "From �����ʿ�Ʒ a,�����ʿ�Ʒ��Ŀ b " & vbNewLine & _
            "Where a.id = b.�ʿ�Ʒid and [1] Between a.��ʼ���� And a.�������� And a.����id = [2]" & vbNewLine & _
            "Order By ��ʼ���� Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, dateWhy, lng����id)
    
    With vsGrid
        .Clear
        .Rows = 2: .Cols = 9
        Set .DataSource = rsTemp
        
        If .Cols > 1 Then
            .ColWidth(0) = 0
            .ColHidden(0) = True
            .ColHidden(1) = True
            .ColHidden(8) = True
            If .Rows > 1 Then
                .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
            End If
            
        End If
        If Not rsTemp.EOF Then .AutoSize 2, .Cols - 1
            
      '  .Select .FixedRows, 1
    End With
    
    Exit Function
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetQCData(ByRef vsGrid As VSFlexGrid, ByVal lngQCID As Long, ByVal dateWhy As Date, strGetQCVal As String) As Boolean
    'ȡQC����
    Dim strsql As String, rsTemp As ADODB.Recordset
    Dim dBegin As Date, dEnd As Date, iCol As Integer, iRow As Integer
    
    On Error GoTo hErr
    
    dBegin = Format(dateWhy, "yyyy-MM-dd 00:00:00")
    dEnd = Format(dateWhy, "yyyy-MM-dd 23:59:59")
    
    strsql = "Select Distinct  F.����, F.������, E.��д, '' as ��һ��,'' as �ڶ���,'' as ������,'' as ���Ĵ�,'' as �����,'' as ������, '' as ���ߴ�,'' as �ڰ˴�,'' as �ھŴ�" & vbNewLine & _
            "       ,A.�ʿ�Ʒid, A.��Ŀid, A.ȡֵ����, A.����ֵ, E.�������,Nvl(G.С��λ��,2) as С��λ��, '' as �걾id,'' as �걾���,'' as ������,'' ���ü�¼,'' as ���" & vbNewLine & _
            "From �����ʿ�Ʒ��Ŀ A, ������Ŀ E, ����������Ŀ F,����������Ŀ G" & vbNewLine & _
            "Where A.��Ŀid = E.������Ŀid And A.��Ŀid = F.ID And A.�ʿ�Ʒid = [1] and A.��Ŀid= G.��ĿID and G.����ID=[2]" & vbNewLine & _
            "Order By F.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, lngQCID, Val(cbo����.ItemData(cbo����.ListIndex)))
    With vsGrid
        .Clear
        .Rows = 2: .Cols = 23
        
        Set .DataSource = rsTemp
        For iCol = 11 To .Cols - 1
            .ColHidden(iCol) = True
        Next
        
        If Not rsTemp.EOF Then .AutoSize 0, 11
 
        'ȡ����
        strsql = "Select f.����, f.������, e.��д, d.������,d.od,d.sco, t.���, e.�������, nvl(i.С��λ��,2) as С��λ��, a.*" & vbNewLine & _
                "From (Select a.�ʿ�Ʒid, a.��Ŀid, c.�걾���, b.�걾id, b.����ʱ��, a.ȡֵ����, a.����ֵ, b.���Դ���, b.������, b.���ü�¼, b.����id" & vbNewLine & _
                "       From �����ʿ�Ʒ��Ŀ A, �����ʿؼ�¼ B, ����걾��¼ C" & vbNewLine & _
                "       Where b.�걾id = c.Id And a.�ʿ�Ʒid = b.�ʿ�Ʒid And a.�ʿ�Ʒid = [1] And" & vbNewLine & _
                "             b.����ʱ�� Between [2] And [3]) A, ������ͨ��� D, ������Ŀ E, ����������Ŀ F, �����ʿر��� T, ����������Ŀ I" & vbNewLine & _
                "Where d.Id = t.���id(+) And a.�걾id = d.����걾id And a.��Ŀid = d.������Ŀid And a.��Ŀid = e.������Ŀid And a.��Ŀid = f.Id And" & vbNewLine & _
                "      a.����id = i.����id And a.��Ŀid = i.��Ŀid"

        Set mrsQCData = zlDatabase.OpenSQLRecord(strsql, Me.Caption, lngQCID, dBegin, dEnd)
        Do Until mrsQCData.EOF
            For iRow = .FixedRows To .Rows - 1
                If .TextMatrix(iRow, 0) = "" & mrsQCData!���� Then
                     
                    If strGetQCVal = "[SCO]" Then
                        .TextMatrix(iRow, 2 + Val("" & mrsQCData!���Դ���)) = Trim("" & mrsQCData!sco)
                        If Val("" & mrsQCData!���) = 2 Then 'ʧ��(��)
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!���Դ���)) = vbRed
                        ElseIf Val("" & mrsQCData!���) = 0 Then '����
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!���Դ���)) = .ForeColor
                        Else  '����(���)
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!���Դ���)) = vbMagenta
                        End If
                        Exit For
                    ElseIf strGetQCVal = "[OD]" Then
                        .TextMatrix(iRow, 2 + Val("" & mrsQCData!���Դ���)) = Trim("" & mrsQCData!od)
                        If Val("" & mrsQCData!���) = 2 Then 'ʧ��(��)
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!���Դ���)) = vbRed
                        ElseIf Val("" & mrsQCData!���) = 0 Then '����
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!���Դ���)) = .ForeColor
                        Else  '����(���)
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!���Դ���)) = vbMagenta
                        End If
                        Exit For
                    Else
                        .TextMatrix(iRow, 2 + Val("" & mrsQCData!���Դ���)) = Trim("" & mrsQCData!������)
                        If Val("" & mrsQCData!���) = 2 Then 'ʧ��(��)
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!���Դ���)) = vbRed
                        ElseIf Val("" & mrsQCData!���) = 0 Then '����
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!���Դ���)) = .ForeColor
                        Else  '����(���)
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!���Դ���)) = vbMagenta
                        End If
                        Exit For
                    End If
                    
                End If
            Next
            mrsQCData.MoveNext
        Loop
    End With
    Exit Function
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function






