VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CO373F~1.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CO1FBF~1.OCX"
Begin VB.Form frmEPRAuditMan 
   Caption         =   "�����������"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   Icon            =   "frmEPRAuditMan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   9615
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picKind 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1020
      Left            =   285
      ScaleHeight     =   1020
      ScaleWidth      =   2325
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   765
      Width           =   2325
      Begin VB.OptionButton optKind 
         Caption         =   "���ﲡ��(&1)"
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   1
         Top             =   75
         Width           =   1380
      End
      Begin VB.OptionButton optKind 
         Caption         =   "סԺ����(&2)"
         Height          =   180
         Index           =   1
         Left            =   420
         TabIndex        =   2
         Top             =   390
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton optKind 
         Caption         =   "������(&3)"
         Height          =   180
         Index           =   2
         Left            =   420
         TabIndex        =   3
         Top             =   720
         Width           =   1380
      End
   End
   Begin VB.PictureBox picDate 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1680
      Left            =   285
      ScaleHeight     =   1680
      ScaleWidth      =   2325
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1935
      Width           =   2325
      Begin VB.CheckBox chkNoData 
         Caption         =   "��ʾ��ҵ�����(&N)"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1425
         Value           =   1  'Checked
         Width           =   1845
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "����ͳ��(&R)"
         Height          =   350
         Left            =   450
         TabIndex        =   6
         Top             =   900
         Width           =   1605
      End
      Begin MSComCtl2.DTPicker dtpDateTo 
         Height          =   300
         Left            =   450
         TabIndex        =   5
         Top             =   465
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   102891523
         CurrentDate     =   38683
      End
      Begin MSComCtl2.DTPicker dtpDateFrom 
         Height          =   300
         Left            =   450
         TabIndex        =   4
         Top             =   120
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   102891523
         CurrentDate     =   38683
      End
      Begin VB.Label lblDateTo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   525
         Width           =   180
      End
      Begin VB.Label lblDateFrom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   180
         Width           =   180
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   6450
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEPRAuditMan.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14076
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
   Begin XtremeSuiteControls.TaskPanel tplThis 
      Height          =   5670
      Left            =   30
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   690
      Width           =   3000
      _Version        =   589884
      _ExtentX        =   5292
      _ExtentY        =   10001
      _StockProps     =   64
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   5310
      Left            =   3090
      TabIndex        =   0
      Top             =   1050
      Width           =   6405
      _cx             =   11298
      _cy             =   9366
      Appearance      =   2
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   3
      FixedRows       =   2
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
      MergeCells      =   1
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
      WallPaper       =   "frmEPRAuditMan.frx":0E1C
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ﲡ��(2005-11-20��2005-11-26)��д���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3090
      TabIndex        =   14
      Top             =   780
      Width           =   4065
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   2325
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmEPRAuditMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mstrPrivs As String     '��ǰʹ����Ȩ�޴�

Private mintKind As Integer     '��������
Private mstrDateFrom As String  '��ʼ����
Private mstrDateTo As String    '��������

'-----------------------------------------------------
'��ʱ����
'-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim rsTemp As New ADODB.Recordset
Dim strSQL As String
Dim lngCount As Long, lngRow As Long, lngCol As Long

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngDeptId As Long, strDeptName As String
    With Me.vfgThis
        lngDeptId = Val(.TextMatrix(.Row, 0))
        strDeptName = .TextMatrix(.Row, 2)
    End With
    
    Select Case Control.ID
    Case conMenu_File_Open:
        Dim cbrPBar As CommandBar
        Dim cbrPItem As CommandBarControl
        
        Set cbrPBar = Me.cbsThis.Add("����", xtpBarPopup)
        With Me.vfgThis
            Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10, "�����ļ��������(&F)")
            If mintKind = 1 Then
                Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 1, "���ﲡ�˲������(&1)")
                cbrPItem.BeginGroup = True
                cbrPItem.Enabled = (Val(.TextMatrix(.Row, 5)) <> 0)
                Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 2, "���ﲡ�˲������(&2)")
                cbrPItem.Enabled = (Val(.TextMatrix(.Row, 7)) <> 0)
            Else
                Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 1, "��Ժ���˲������(&1)")
                cbrPItem.BeginGroup = True
                cbrPItem.Enabled = (Val(.TextMatrix(.Row, 5)) <> 0)
                Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 2, "ת�벡�˲������(&2)")
                cbrPItem.Enabled = (Val(.TextMatrix(.Row, 7)) <> 0)
                Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 3, "��Ժ���˲������(&3)")
                cbrPItem.Enabled = (Val(.TextMatrix(.Row, 9)) <> 0)
                Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 4, "�������˲������(&4)")
                cbrPItem.Enabled = (Val(.TextMatrix(.Row, 11)) <> 0)
                Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 5, "ת�����˲������(&5)")
                cbrPItem.Enabled = (Val(.TextMatrix(.Row, 13)) <> 0)
                If mintKind = 2 Then
                    Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 6, "�����������(&6)")
                    cbrPItem.Enabled = (Val(.TextMatrix(.Row, 15)) <> 0)
                End If
            End If
            Set cbrPItem = cbrPBar.Controls.Add(xtpControlButton, conMenu_File_Open * 10 + 9, "ȫ�岡�˲������(&A)")
            cbrPItem.BeginGroup = True
        End With
        cbrPBar.ShowPopup
    Case conMenu_File_Open * 10: Call frmEPRAuditFile.ShowMe(Me, lngDeptId, strDeptName, mintKind, mstrDateFrom, mstrDateTo)
    Case conMenu_File_Open * 10 + 1: Call frmEPRAuditPati.ShowMe(Me, lngDeptId, strDeptName, mintKind, mstrDateFrom, mstrDateTo, IIf(mintKind = 1, "����", "��Ժ"))
    Case conMenu_File_Open * 10 + 2: Call frmEPRAuditPati.ShowMe(Me, lngDeptId, strDeptName, mintKind, mstrDateFrom, mstrDateTo, IIf(mintKind = 1, "����", "ת��"))
    Case conMenu_File_Open * 10 + 3: Call frmEPRAuditPati.ShowMe(Me, lngDeptId, strDeptName, mintKind, mstrDateFrom, mstrDateTo, "��Ժ")
    Case conMenu_File_Open * 10 + 4: Call frmEPRAuditPati.ShowMe(Me, lngDeptId, strDeptName, mintKind, mstrDateFrom, mstrDateTo, "����")
    Case conMenu_File_Open * 10 + 5: Call frmEPRAuditPati.ShowMe(Me, lngDeptId, strDeptName, mintKind, mstrDateFrom, mstrDateTo, "ת��")
    Case conMenu_File_Open * 10 + 6: Call frmEPRAuditPati.ShowMe(Me, lngDeptId, strDeptName, mintKind, mstrDateFrom, mstrDateTo, "����")
    Case conMenu_File_Open * 10 + 9: Call frmEPRAuditPati.ShowMe(Me, lngDeptId, strDeptName, mintKind, mstrDateFrom, mstrDateTo)
    
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview:  Call zlRptPrint(0)
    Case conMenu_File_Print:    Call zlRptPrint(1)
    Case conMenu_File_Excel:    Call zlRptPrint(3)
    Case conMenu_File_Exit:     Unload Me
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh: Call RefreshData
    Case conMenu_View_Jump
        If Me.Visible Then Me.vfgThis.SetFocus
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    Case Else
        'ִ�з�������ǰģ��ı���
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                "��ʼ����=" & Format(dtpDateFrom.Value, "yyyy-MM-dd"), "��������=" & Format(dtpDateTo.Value, "yyyy-MM-dd"), _
                "��������=" & IIf(optKind(0).Value, "���ﲡ��", IIf(optKind(1).Value, "סԺ����", "������")) & "|" & IIf(optKind(0).Value, 1, IIf(optKind(1).Value, 2, 4)))
        End If
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop  As Long, lngScaleRight  As Long, lngScaleBottom  As Long
    Call Me.cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    With Me.tplThis
        .Left = lngScaleLeft
        .Top = lngScaleTop: .Height = lngScaleBottom - .Top
    End With
    With Me.lblTitle
        .Left = Me.tplThis.Left + Me.tplThis.Width + 30: .Width = lngScaleRight - .Left
        .Top = lngScaleTop + 60
    End With
    With Me.vfgThis
        .Left = Me.tplThis.Left + Me.tplThis.Width: .Width = lngScaleRight - .Left
        .Top = Me.lblTitle.Top + Me.lblTitle.Height + 60: .Height = lngScaleBottom - .Top
    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Open
        With Me.vfgThis
            Control.Enabled = (Val(.TextMatrix(.Row, 1)) <> 0)
            If Control.Enabled = False Then Exit Sub
            For lngCol = 3 To .Cols - 1
                Control.Enabled = (Val(.TextMatrix(.Row, lngCol)) <> 0)
                If Control.Enabled Then Exit Sub
            Next
        End With
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.vfgThis.Rows > Me.vfgThis.FixedRows + 1)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub chkNoData_Click()
    Dim blnData As Boolean
    With Me.vfgThis
        If Me.chkNoData.Value = vbChecked Then
            For lngRow = .FixedRows To .Rows - 2
                .ROWHEIGHT(lngRow) = .RowHeightMin
                .RowHidden(lngRow) = False
            Next
        Else
            For lngRow = .FixedRows To .Rows - 2
                blnData = False
                For lngCol = 3 To .Cols - 1
                    If Val(.TextMatrix(lngRow, lngCol)) <> 0 Then blnData = True: Exit For
                Next
                If blnData = False Then
                    .ROWHEIGHT(lngRow) = 0
                    .RowHidden(lngRow) = True
                End If
            Next
        End If
    End With
End Sub

Private Sub cmdSearch_Click()
    If Me.dtpDateTo.Value - Me.dtpDateFrom.Value > 15 Then MsgBox "���ʱ�䷶Χ̫��(���ܳ���15��)��", vbExclamation, gstrSysName: Exit Sub
    
    If Me.optKind(0).Value Then
        mintKind = 1
    ElseIf Me.optKind(1).Value Then
        mintKind = 2
    ElseIf Me.optKind(2).Value Then
        mintKind = 4
    Else
        Me.optKind(1).Value = True: mintKind = 2
    End If
    mstrDateFrom = Format(Me.dtpDateFrom.Value, "yyyy-mm-dd")
    mstrDateTo = Format(Me.dtpDateTo.Value, "yyyy-mm-dd")
    
    Call RefreshData
End Sub

Private Sub dtpDateFrom_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub dtpDateTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub dtpDateTo_Validate(Cancel As Boolean)
    Me.dtpDateFrom.MaxDate = Me.dtpDateTo.Value
    If Me.dtpDateFrom.Value > Me.dtpDateFrom.MaxDate Then Me.dtpDateFrom.Value = Me.dtpDateFrom.MaxDate
End Sub

Private Sub Form_Load()
    Call zlcommfun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
    mstrPrivs = gstrPrivs
    
    '-----------------------------------------------------
    '��ʼ����
    Call InitTerm
    Call cmdSearch_Click
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = frmPubIcons.imgPublic.Icons
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
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "չ��(&O)��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Jump, "��ת�����(&J)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("O"), conMenu_File_Open
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F6, conMenu_View_Jump
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Jump
    End With
    
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "չ��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '��ȡ��������ģ��ı���:��Ϊ��һ���Զ�ȡ,ȫ�ֱ�������
    '---------------------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    '-----------------------------------------------------
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ��ҵ�����", IIf(Me.chkNoData.Value = vbChecked, 1, 0))
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub optKind_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub vfgThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbRightButton Then Exit Sub

    Set cbrControl = Me.cbsThis.FindControl(, conMenu_File_Open)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub InitTerm()
    '-------------------------------------------------
    '--���ܣ���ʼ�������Ͳ���
    '-------------------------------------------------
    Dim tplGroup As TaskPanelGroup
    Dim tplItem As TaskPanelGroupItem
    
    '-----------------------------------------------------
    '��ʼ����:
    If GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ��ҵ�����", 1) = 1 Then
        Me.chkNoData.Value = vbChecked
    Else
        Me.chkNoData.Value = vbUnchecked
    End If
    
    strSQL = "Select Sysdate From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With Me.dtpDateTo
        .Value = Format(rsTemp.Fields(0).Value, "yyyy-MM-dd")
        .MaxDate = .Value: .MinDate = Format("1990-01-01", "yyyy-MM-dd")
    End With
    With Me.dtpDateFrom
        .Value = Me.dtpDateTo.Value - 7
        .MaxDate = Me.dtpDateTo.MaxDate: .MinDate = Me.dtpDateTo.MinDate
    End With
    
    '-----------------------------------------------------
    '��ʾ��̬
    Set tplGroup = Me.tplThis.Groups.Add(0, "��鷶Χ:"): tplGroup.Expandable = False
    Set tplItem = tplGroup.Items.Add(0, "��������:", xtpTaskItemTypeText)
    Set tplItem = tplGroup.Items.Add(0, "", xtpTaskItemTypeControl): Set tplItem.Control = Me.picKind
    Me.picKind.BackColor = tplItem.BackColor
    For lngCount = 0 To Me.optKind.Count - 1: Me.optKind(lngCount).BackColor = tplItem.BackColor: Next
    Set tplItem = tplGroup.Items.Add(0, "��д���ڷ�Χ:", xtpTaskItemTypeText)
    Set tplItem = tplGroup.Items.Add(0, "", xtpTaskItemTypeControl): Set tplItem.Control = Me.picDate
    Me.picDate.BackColor = tplItem.BackColor
    Me.chkNoData.BackColor = tplItem.BackColor
    
    Set tplGroup = Me.tplThis.Groups.Add(0, "����˵��:"): tplGroup.Expandable = True
    Set tplItem = tplGroup.Items.Add(0, "  �١���Ӧ���¼�����ɲ�������������Ҫ����дһ�εĲ�����������Ҫ��ѭ����д�Ĳ�����", xtpTaskItemTypeText)
    Set tplItem = tplGroup.Items.Add(0, "  �ڡ�����ĳЩ�¼���ӦҪ����дһ�εĲ���Ϊ���֣��������ɲ��������ܳ��������˴�����", xtpTaskItemTypeText)

    '-----------------------------------------------------
    Me.tplThis.Reposition
    Me.BackColor = tplItem.BackColor
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub RefreshData()
    '-------------------------------------------------
    '����:������鷶Χ��֯��ʾ�������
    '-------------------------------------------------
    Dim lngTotal As Long
    
    Select Case mintKind
    Case 1  '���ﲡ��
        Me.lblTitle.Caption = "���ﲡ��(" & mstrDateFrom & "��" & mstrDateTo & ")��д���"
        strSQL = "Select D.ID, D.����, D.����, W.�����, W.����д, P.�����˴�, W.�������, P.�����˴�, W.�������" & vbNewLine & _
                " From ���ű� D, ��������˵�� M," & vbNewLine & _
                "      (Select ִ�в���id, Sum(Decode(����, 1, 0, 1)) As �����˴�, Sum(Decode(����, 1, 1, 0)) As �����˴�" & vbNewLine & _
                "        From ���˹Һż�¼" & vbNewLine & _
                "        Where Nvl(ִ��״̬, 0) <> 0 And �Ǽ�ʱ�� Between To_Date([1], 'yyyy-mm-dd') And" & vbNewLine & _
                "              To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ִ�в���id) P," & vbNewLine & _
                "      (Select W.����id, Sum(W.�����) As �����, Sum(W.����д) As ����д," & vbNewLine & _
                "               Sum(Decode(F.�¼�, '����', W.�����, Null)) As �������," & vbNewLine & _
                "               Sum(Decode(F.�¼�, '����', W.�����, Null)) As �������" & vbNewLine & _
                "        From (Select F.ID, F.ͨ��, A.����id, Q.�¼�" & vbNewLine & _
                "               From �����ļ��б� F, ����Ӧ�ÿ��� A, ����ʱ��Ҫ�� Q" & vbNewLine & _
                "               Where F.ID = A.�ļ�id(+) And F.ID = Q.�ļ�id And F.���� = 1) F," & vbNewLine & _
                "             (Select ����id, �ļ�id, Sum(Decode(���ʱ��, Null, 0, 1)) As �����," & vbNewLine & _
                "                      Sum(Decode(���ʱ��, Null, 1, 0)) As ����д" & vbNewLine & _
                "               From ���Ӳ�����¼" & vbNewLine & _
                "               Where �������� = 1 And ����ʱ�� Between To_Date([1], 'yyyy-mm-dd') And" & vbNewLine & _
                "                     To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "               Group By ����id, �ļ�id) W" & vbNewLine & _
                "        Where F.ID = W.�ļ�id And (F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = W.����id)" & vbNewLine & _
                "        Group By W.����id) W" & vbNewLine & _
                " Where D.ID = M.����id And M.�������� = '�ٴ�' And M.������� In (1, 3) And D.ID = P.ִ�в���id(+) And" & vbNewLine & _
                "       D.ID = W.����id(+)" & vbNewLine & _
                " Order By D.����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrDateFrom, mstrDateTo)
        
        With Me.vfgThis
            .Clear
            Set .DataSource = rsTemp
            
            .ColWidth(0) = 0: .ColHidden(0) = True
            .MergeCells = flexMergeFree
            .MergeRow(0) = True
            .TextMatrix(0, 1) = "����": .TextMatrix(0, 2) = .TextMatrix(0, 1)
            .TextMatrix(0, 3) = "������д���": .TextMatrix(0, 4) = .TextMatrix(0, 3)
            .TextMatrix(0, 5) = "����": .TextMatrix(0, 6) = .TextMatrix(0, 5)
            .TextMatrix(0, 7) = "����": .TextMatrix(0, 8) = .TextMatrix(0, 7)
            
            .TextMatrix(1, 1) = "����": .TextMatrix(1, 2) = "����"
            .TextMatrix(1, 3) = "�����": .TextMatrix(1, 4) = "����д"
            .TextMatrix(1, 5) = "�˴�": .TextMatrix(1, 6) = "��ɲ���"
            .TextMatrix(1, 7) = "�˴�": .TextMatrix(1, 8) = "��ɲ���"
        End With
    
    Case 2  'סԺ����
        Me.lblTitle.Caption = "סԺ����(" & mstrDateFrom & "��" & mstrDateTo & ")��д���"
        strSQL = "Select D.ID, D.����, D.����, W.�����, W.����д, I.��Ժ�˴�, W.��Ժ����, E.ת���˴�, W.ת�벡��, O.��Ժ�˴�," & vbNewLine & _
                "        W.��Ժ����, O.�����˴�, W.��������, G.ת���˴�, W.ת������, S.�����˴�, W.��������" & vbNewLine & _
                " From ���ű� D, ��������˵�� M," & vbNewLine & _
                "      (Select W.����id, Sum(�����) As �����, Sum(����д) As ����д," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, '��Ժ', 1, '�״���Ժ', 1, '�ٴ���Ժ', 1, 0), 0) * �����) As ��Ժ����," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, 'ת��', Decode(Sign(F.��дʱ��), -1, 0, 1), 0), 0) * �����) As ת�벡��," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, '��Ժ', 1, '24Сʱ��Ժ', 1, 0), 0) * �����) As ��Ժ����," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, '����', 1, '24Сʱ����', 1, 0), 0) * �����) As ��������," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, 'ת��', Decode(Sign(F.��дʱ��), -1, 1, 0), 0), 0) * �����) As ת������," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, '����', 1, 0), 0) * �����) As ��������" & vbNewLine & _
                "        From (Select F.ID, F.ͨ��, A.����id, Q.�¼�, Q.Ψһ, Q.��дʱ��" & vbNewLine & _
                "               From �����ļ��б� F, ����Ӧ�ÿ��� A, ����ʱ��Ҫ�� Q" & vbNewLine & _
                "               Where F.ID = A.�ļ�id(+) And F.ID = Q.�ļ�id And F.���� = 2) F," & vbNewLine & _
                "             (Select ����id, �ļ�id, Sum(Decode(���ʱ��, Null, 0, 1)) As �����," & vbNewLine & _
                "                      Sum(Decode(���ʱ��, Null, 1, 0)) As ����д" & vbNewLine & _
                "               From ���Ӳ�����¼" & vbNewLine & _
                "               Where �������� = 2 And ����ʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "               Group By ����id, �ļ�id) W" & vbNewLine & _
                "        Where F.ID = W.�ļ�id And (F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = W.����id)" & vbNewLine & _
                "        Group By W.����id) W," & vbNewLine
        strSQL = strSQL & "      (Select ��Ժ����id, Count(*) As ��Ժ�˴�" & vbNewLine & _
                "        From ������ҳ" & vbNewLine & _
                "        Where ��Ժ���� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ��Ժ����id) I," & vbNewLine & _
                "      (Select ����id, Count(*) As ת���˴�" & vbNewLine & _
                "        From ���˱䶯��¼" & vbNewLine & _
                "        Where ��ʼԭ�� = 3 And ��ʼʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ����id) E," & vbNewLine & _
                "      (Select ��Ժ����id, Sum(Decode(��Ժ��ʽ, '����', 0, 1)) As ��Ժ�˴�, Sum(Decode(��Ժ��ʽ, '����', 1, 0)) As �����˴�" & vbNewLine & _
                "        From ������ҳ" & vbNewLine & _
                "        Where ��Ժ���� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ��Ժ����id) O," & vbNewLine & _
                "      (Select ����id, Count(*) As ת���˴�" & vbNewLine & _
                "        From ���˱䶯��¼" & vbNewLine & _
                "        Where ��ֹԭ�� = 3 And ��ֹʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ����id) G," & vbNewLine & _
                "      (Select R.���˿���id, Count(*) As �����˴�" & vbNewLine & _
                "        From ����ҽ����¼ R, ����ҽ������ S" & vbNewLine & _
                "        Where R.ID = S.ҽ��id And R.������� = 'F' And S.�״�ʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By R.���˿���id) S" & vbNewLine & _
                " Where D.ID = M.����id And M.�������� = '�ٴ�' And ������� In (2, 3) And D.ID = W.����id(+) And D.ID = I.��Ժ����id(+) And" & vbNewLine & _
                "       D.ID = E.����id(+) And D.ID = O.��Ժ����id(+) And D.ID = G.����id(+) And D.ID = S.���˿���id(+)" & vbNewLine & _
                " Order By D.����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrDateFrom, mstrDateTo)
        
        With Me.vfgThis
            .Clear
            Set .DataSource = rsTemp
            
            .ColWidth(0) = 0: .ColHidden(0) = True
            .MergeCells = flexMergeFree
            .MergeRow(0) = True
            .TextMatrix(0, 1) = "����": .TextMatrix(0, 2) = .TextMatrix(0, 1)
            .TextMatrix(0, 3) = "������д���": .TextMatrix(0, 4) = .TextMatrix(0, 3)
            .TextMatrix(0, 5) = "��Ժ": .TextMatrix(0, 6) = .TextMatrix(0, 5)
            .TextMatrix(0, 7) = "ת��": .TextMatrix(0, 8) = .TextMatrix(0, 7)
            .TextMatrix(0, 9) = "��Ժ": .TextMatrix(0, 10) = .TextMatrix(0, 9)
            .TextMatrix(0, 11) = "����": .TextMatrix(0, 12) = .TextMatrix(0, 11)
            .TextMatrix(0, 13) = "ת��": .TextMatrix(0, 14) = .TextMatrix(0, 13)
            .TextMatrix(0, 15) = "����": .TextMatrix(0, 16) = .TextMatrix(0, 15)
            
            .TextMatrix(1, 1) = "����": .TextMatrix(1, 2) = "����"
            .TextMatrix(1, 3) = "�����": .TextMatrix(1, 4) = "����д"
            .TextMatrix(1, 5) = "�˴�": .TextMatrix(1, 6) = "��ɲ���"
            .TextMatrix(1, 7) = "�˴�": .TextMatrix(1, 8) = "��ɲ���"
            .TextMatrix(1, 9) = "�˴�": .TextMatrix(1, 10) = "��ɲ���"
            .TextMatrix(1, 11) = "�˴�": .TextMatrix(1, 12) = "��ɲ���"
            .TextMatrix(1, 13) = "�˴�": .TextMatrix(1, 14) = "��ɲ���"
            .TextMatrix(1, 15) = "�˴�": .TextMatrix(1, 16) = "��ɲ���"
        End With
    Case 4  '������
        Me.lblTitle.Caption = "������(" & mstrDateFrom & "��" & mstrDateTo & ")��д���"
        strSQL = "Select D.ID, D.����, D.����, W.�����, W.����д, I.��Ժ�˴�, W.��Ժ����, E.ת���˴�, W.ת�벡��, O.��Ժ�˴�," & vbNewLine & _
                "        W.��Ժ����, O.�����˴�, W.��������, G.ת���˴�, W.ת������" & vbNewLine & _
                " From ���ű� D, ��������˵�� M," & vbNewLine & _
                "      (Select W.����id, Sum(�����) As �����, Sum(����д) As ����д," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, '��Ժ', 1, '�״���Ժ', 1, '�ٴ���Ժ', 1, 0), 0) * �����) As ��Ժ����," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, 'ת��', Decode(Sign(F.��дʱ��), -1, 0, 1), 0), 0) * �����) As ת�벡��," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, '��Ժ', 1, '24Сʱ��Ժ', 1, 0), 0) * �����) As ��Ժ����," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, '����', 1, '24Сʱ����', 1, 0), 0) * �����) As ��������," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, 'ת��', Decode(Sign(F.��дʱ��), -1, 1, 0), 0), 0) * �����) As ת������" & vbNewLine & _
                "        From (Select F.ID, F.ͨ��, A.����id, Q.�¼�, Q.Ψһ, Q.��дʱ��" & vbNewLine & _
                "               From �����ļ��б� F, ����Ӧ�ÿ��� A, ����ʱ��Ҫ�� Q" & vbNewLine & _
                "               Where F.ID = A.�ļ�id(+) And F.ID = Q.�ļ�id And F.���� = 4) F," & vbNewLine & _
                "             (Select ����id, �ļ�id, Sum(Decode(���ʱ��, Null, 0, 1)) As �����," & vbNewLine & _
                "                      Sum(Decode(���ʱ��, Null, 1, 0)) As ����д" & vbNewLine & _
                "               From ���Ӳ�����¼" & vbNewLine & _
                "               Where �������� = 4 And ����ʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "               Group By ����id, �ļ�id) W" & vbNewLine & _
                "        Where F.ID = W.�ļ�id And (F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = W.����id)" & vbNewLine & _
                "        Group By W.����id) W," & vbNewLine
        strSQL = strSQL & "      (Select ��Ժ����id, Count(*) As ��Ժ�˴�" & vbNewLine & _
                "        From ������ҳ" & vbNewLine & _
                "        Where ��Ժ���� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ��Ժ����id) I," & vbNewLine & _
                "      (Select ����id, Count(*) As ת���˴�" & vbNewLine & _
                "        From ���˱䶯��¼" & vbNewLine & _
                "        Where ��ʼԭ�� = 3 And ��ʼʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ����id) E," & vbNewLine & _
                "      (Select ��ǰ����id, Sum(Decode(��Ժ��ʽ, '����', 0, 1)) As ��Ժ�˴�, Sum(Decode(��Ժ��ʽ, '����', 1, 0)) As �����˴�" & vbNewLine & _
                "        From ������ҳ" & vbNewLine & _
                "        Where ��Ժ���� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ��ǰ����id) O," & vbNewLine & _
                "      (Select ����id, Count(*) As ת���˴�" & vbNewLine & _
                "        From ���˱䶯��¼" & vbNewLine & _
                "        Where ��ֹԭ�� = 3 And ��ֹʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ����id) G" & vbNewLine & _
                " Where D.ID = M.����id And M.�������� = '����' And ������� In (2, 3) And D.ID = W.����id(+) And D.ID = I.��Ժ����id(+) And" & vbNewLine & _
                "       D.ID = E.����id(+) And D.ID = O.��ǰ����id(+) And D.ID = G.����id(+)" & vbNewLine & _
                " Order By D.����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrDateFrom, mstrDateTo)
        
        With Me.vfgThis
            .Clear
            Set .DataSource = rsTemp
            
            .ColWidth(0) = 0: .ColHidden(0) = True
            .MergeCells = flexMergeFree
            .MergeRow(0) = True
            .TextMatrix(0, 1) = "����": .TextMatrix(0, 2) = .TextMatrix(0, 1)
            .TextMatrix(0, 3) = "������д���": .TextMatrix(0, 4) = .TextMatrix(0, 3)
            .TextMatrix(0, 5) = "��Ժ": .TextMatrix(0, 6) = .TextMatrix(0, 5)
            .TextMatrix(0, 7) = "ת��": .TextMatrix(0, 8) = .TextMatrix(0, 7)
            .TextMatrix(0, 9) = "��Ժ": .TextMatrix(0, 10) = .TextMatrix(0, 9)
            .TextMatrix(0, 11) = "����": .TextMatrix(0, 12) = .TextMatrix(0, 11)
            .TextMatrix(0, 13) = "ת��": .TextMatrix(0, 14) = .TextMatrix(0, 13)
            
            .TextMatrix(1, 1) = "����": .TextMatrix(1, 2) = "����"
            .TextMatrix(1, 3) = "�����": .TextMatrix(1, 4) = "����д"
            .TextMatrix(1, 5) = "�˴�": .TextMatrix(1, 6) = "��ɲ���"
            .TextMatrix(1, 7) = "�˴�": .TextMatrix(1, 8) = "��ɲ���"
            .TextMatrix(1, 9) = "�˴�": .TextMatrix(1, 10) = "��ɲ���"
            .TextMatrix(1, 11) = "�˴�": .TextMatrix(1, 12) = "��ɲ���"
            .TextMatrix(1, 13) = "�˴�": .TextMatrix(1, 14) = "��ɲ���"
        End With
    End Select
    
    '��ϼ�
    With Me.vfgThis
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 2) = "�ϼ�"
        For lngCol = 3 To .Cols - 1
            lngTotal = 0
            For lngRow = .FixedRows To .Rows - 2
                lngTotal = lngTotal + Val(.TextMatrix(lngRow, lngCol))
            Next
            .TextMatrix(.Rows - 1, lngCol) = lngTotal
        Next
        .Row = .FixedRows: .Col = 1
        Call .AutoSize(1, .Cols - 1)
    End With
    
    '��ʾ�����ؿ���
    Call chkNoData_Click
    Me.stbThis.Panels(2).Text = "�����չ��(Ctrl+O)����ϸ��鵱ǰ���Ҳ��˲����������������д�����"
    
    If Me.Visible Then Me.vfgThis.SetFocus
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgThis
    objPrint.Title.Text = Me.lblTitle.Caption
    Set objPrint.Title.Font = Me.lblTitle.Font
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objPrint.UnderAppRows.Add(objAppRow)
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    Me.vfgThis.Tag = "Printing"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Me.vfgThis.Tag = ""
End Sub

