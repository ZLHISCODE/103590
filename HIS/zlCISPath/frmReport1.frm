VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReport1 
   AutoRedraw      =   -1  'True
   Caption         =   "��������ط�������ָ��������"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   Icon            =   "frmReport1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11835
   Begin VB.ComboBox cboDate 
      Height          =   300
      Left            =   9720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1320
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8130
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmReport1.frx":6852
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17965
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
   Begin VB.Frame fraReport 
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   11655
      Begin VSFlex8Ctl.VSFlexGrid vsItem 
         Height          =   7530
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   11535
         _cx             =   20346
         _cy             =   13282
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14744288
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
         RowHeightMax    =   350
         ColWidthMin     =   0
         ColWidthMax     =   8000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmReport1.frx":70E6
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         WordWrap        =   -1  'True
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
         BackColorFrozen =   14811105
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   240
      Top             =   180
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmReport1.frx":71A6
      Left            =   825
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmReport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String

Private mstr·������ As String
Private mlng·��ID  As Long
Private mblnEdit As Boolean '��ǰ�Ƿ��ڱ༭ģʽ
Private mstr��ǰ�ڼ� As String '��ǰѡ����ڼ�
Private mstrǰһ�ڼ� As String
Private mrsDate As ADODB.Recordset '�ڼ�
Private mrsPati As ADODB.Recordset 'ĳ��������ָ����Ժʱ�䷶Χ�ڵĲ���

Private Const conMenu_Date = 400
Private Const conMenu_Edit_GetCur = 3052
Private Const conMenu_Edit_GetAll = 3013

Private Enum ����
    COL��� = 0
    COL��Ŀ�ı�1 = 1
    COL��Ŀ�ı�2 = 2
    col��� = 3
    col��ע = 4
End Enum


Public Function ShowMe(frmMain As Object, ByVal lng·��ID As Long, ByVal str·������ As String) As Boolean
    mstr·������ = str·������
    mlng·��ID = lng·��ID
    
    Me.Show 1, frmMain
End Function

Private Sub cboDate_Click()
    
    If cboDate.ListIndex >= 0 Then
        mrsDate.Filter = "ID=" & cboDate.ItemData(cboDate.ListIndex)
        Me.dkpMain.Panes(1).Title = "ʱ�䷶Χ��" & Format(mrsDate!��ʼʱ��, "yyyy��mm��dd��") & _
                    "��" & Format(mrsDate!����ʱ��, "yyyy��mm��dd��") & "        ���֣�" & mstr·������
        mstrǰһ�ڼ� = mstr��ǰ�ڼ�
        mstr��ǰ�ڼ� = mrsDate!�ڼ�
        Call zlRefresh
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And mblnEdit Then
        Call ExeCancelEdit
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Or KeyAscii = Asc("|") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    mblnEdit = False
    fraReport.Visible = False
    mstrPrivs = gstrPrivs
    
    Call InitCommandBar
    Call InitDockPannel
      
    Call InitListTable
    
    Call FillStructure
    Call FillDate '��䱨������
    
    Call RestoreWinState(Me, App.ProductName)
        
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    mstr��ǰ�ڼ� = ""
    mstrǰһ�ڼ� = ""
    Set mrsDate = Nothing
    Set mrsPati = Nothing
End Sub

Private Sub InitCommandBar()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustomControl As CommandBarControlCustom

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
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
    
    '�˵�����:������������
    '    ���xtpControlPopup���͵�����ID���¸�ֵ
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_GetAll, "��ȡ��������(&R)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_GetCur, "��ȡ��ǰ������(&X)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend, "�༭SQL(&E)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True
    End With
    
    '���˵��Ҳ�Ĳ���
    With cbsMain.ActiveMenuBar.Controls
        Set objControl = .Add(xtpControlLabel, 0, "�ڼ� ")
        objControl.Flags = xtpFlagRightAlign
        Set objCustomControl = .Add(xtpControlCustom, conMenu_Date, "")
        objCustomControl.Handle = cboDate.Hwnd
        objCustomControl.Flags = xtpFlagRightAlign
    End With

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_GetAll, "��ȡ��������")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_GetCur, "��ȡ��ǰ������")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend, "�༭SQL")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "����")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��")
                        
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each objControl In objBar.Controls
        If objControl.Type = xtpControlButton Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    '����Ŀ����:���������������Ѵ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '��ӡ
        
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '����
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '�޸�
        
        .Add 0, vbKeyF6, conMenu_Edit_GetCur '��ȡ��ǰ��
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        .Add 0, vbKeyF4, conMenu_Edit_Compend '�༭SQL
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save '����
        .Add FCONTROL, vbKeyS, conMenu_Edit_Transf_Save '����
        
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With
    
End Sub

Private Sub InitDockPannel()
    Dim objPane As Pane

    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True

    Set objPane = Me.dkpMain.CreatePane(1, 600, 600, DockTopOf, Nothing)
    objPane.Title = "���֣�" & mstr·������
    objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    objPane.MinTrackSize.SetSize Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub InitListTable()
'���ܣ���ʼ�������嵥���
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "���,1000,4;����ָ��,1800,4;����ָ��,3200,4;���,1200,7;��ע,2000,7"
    arrHead = Split(strHead, ";")
    With vsItem
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 'Ϊ��֧��zl9PrintMode
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
    End With
End Sub
'
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = fraReport.Hwnd
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    
'    '���������ؼ�Resize����
    vsItem.Width = fraReport.Width
    vsItem.Height = fraReport.Height
End Sub


Private Sub FuncPathTableOutput(bytStyle As Byte)
'���ܣ��������
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim rsTmp As ADODB.Recordset
    
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim lngRow As Long, lngCol As Long
    Dim lngColor As Long, bytR As Byte
        
    '��ͷ
    objOut.Title.Text = Me.Caption
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    Set objRow = New zlTabAppRow
    objRow.Add " "
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "                                              ���֣�" & mstr·������
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    If mstr��ǰ�ڼ� <> "" Then
        objRow.Add "ʱ�䷶Χ��" & Format(mrsDate!��ʼʱ��, "yyyy��mm��dd��") & _
                        "��" & Format(mrsDate!����ʱ��, "yyyy��mm��dd��")
    End If
    objOut.UnderAppRows.Add objRow
    
    '����
'    Set objRow = New zlTabAppRow
'    objRow.Add "��ӡ�ˣ�" & UserInfo.����
'    objRow.Add "��ӡ���ڣ�" & Format(zldatabase.Currentdate(), "yyyy��MM��dd��")
'    objOut.BelowAppRows.Add objRow
    
    '����
    Set objOut.Body = vsItem
    
    '���
    With vsItem
        If bytStyle = 1 Then
            bytR = zlPrintAsk(objOut)
            Me.Refresh
            If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
        Else
            zlPrintOrView1Grd objOut, bytStyle
        End If
    End With
   
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    
    Select Case Control.ID
     '0.���
    Case conMenu_File_PrintSet
        Call zlPrintSet
    Case conMenu_File_Print
        Call FuncPathTableOutput(1)
    Case conMenu_File_Preview
        Call FuncPathTableOutput(2)
    Case conMenu_File_Excel
        Call FuncPathTableOutput(3)
            
    Case conMenu_Edit_NewItem '����
        Call ExeNew
        
    Case conMenu_Edit_Modify '�޸�
        Call ExeModify
        
    Case conMenu_Edit_Delete 'ɾ��
        Call ExeDelete
   
    Case conMenu_Edit_GetAll    '��ȡ��������
        Call ExeGetData
    Case conMenu_Edit_GetCur    '��ȡ��ǰ������
        If vsItem.Row < vsItem.FixedRows Then
            Me.stbThis.Panels(2) = "����ѡ�б����е�һ�С�"
            Exit Sub
        End If
        Call ExeGetData(vsItem.Row)
    Case conMenu_Edit_Compend   '�༭SQL
        If vsItem.Row < vsItem.FixedRows Then
            Me.stbThis.Panels(2) = "����ѡ�б����е�һ�С�"
            Exit Sub
        End If
        
        Call ExeDefineSQL
        
    Case conMenu_Edit_Transf_Save   '����
        Call ExeSaveData
    Case conMenu_Edit_Transf_Cancle 'ȡ��
        Call ExeCancelEdit
   
    Case conMenu_View_ToolBar_Button '������
        For i = 2 To cbsMain.count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '��ť����
        For i = 2 To cbsMain.count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = _
                IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '��ͼ��
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '״̬��
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
  
    Case conMenu_View_Refresh 'ˢ��
        Call FillStructure
        Call zlRefresh
        Me.stbThis.Panels(2).Text = "�����ɹ���"
   
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.Hwnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.Hwnd)
    Case conMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '�˳�
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    
    If mblnEdit Then    '�༭״̬
        Select Case Control.ID
            Case conMenu_EditPopup, conMenu_HelpPopup, conMenu_Edit_GetAll, conMenu_Edit_GetCur, conMenu_Edit_Transf_Save, _
                conMenu_Edit_Transf_Cancle, conMenu_Edit_Compend  '��ȡ����,����,ȡ��,��������Դ
                Control.Enabled = True
            Case Else
                Control.Enabled = False
        End Select
    Else
        
        '����Ȩ�����ð�ť�ɼ�״̬
        Call SetControlVisible(Control)
        If Not Control.Visible Then Exit Sub
    
        Select Case Control.ID
            Case conMenu_Edit_Modify, conMenu_Edit_Delete
                Control.Enabled = mstr��ǰ�ڼ� <> ""
                
            Case conMenu_Edit_GetAll, conMenu_Edit_GetCur, conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle '��ȡ����,����,ȡ��
                Control.Enabled = False
                 
            Case conMenu_View_ToolBar_Button '������
                If cbsMain.count >= 2 Then
                    Control.Checked = Me.cbsMain(2).Visible
                End If
            Case conMenu_View_ToolBar_Text 'ͼ������
                If cbsMain.count >= 2 Then
                    Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
                End If
            Case conMenu_View_ToolBar_Size '��ͼ��
                Control.Checked = Me.cbsMain.Options.LargeIcons
            Case conMenu_View_StatusBar '״̬��
                Control.Checked = Me.stbThis.Visible
            Case Else
                Control.Enabled = True
        End Select
    End If
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'���ܣ�����Ȩ�����ò˵��͹������Ŀɼ�״̬
    Dim blnVisible As Boolean, strItem As String

    'Ȩ��ֻ���ж�һ��,�Ѿ��жϹ�����������ж�
    If Control.Category = "���ж�" Then Exit Sub

    blnVisible = True
    Select Case Control.ID
        Case conMenu_Edit_Compend
            If InStr(";" & mstrPrivs & ";", ";�����ֱ�����;") = 0 Then blnVisible = False
    End Select
    
    Control.Visible = blnVisible
    Control.Category = "���ж�"
End Sub

Private Sub SetVsitemEdit(ByVal blnEnable As Boolean)
'���ܣ����ñ��Ŀ�����״̬
'������bytEnable:=0����������,1=��������
    
    If blnEnable Then
        vsItem.Editable = flexEDKbdMouse
        mblnEdit = True
    Else
        vsItem.Editable = flexEDNone
        mblnEdit = False
    End If
End Sub

Private Sub ExeCancelEdit()
    
    If MsgBox("���������󣬱��α༭�����ݽ����ᱻ���棬��ȷ��Ҫ������", vbInformation + vbYesNo, gstrSysName) = vbNo Then
        Exit Sub
    End If
    Call SetVsitemEdit(False)
    Call zlRefresh
End Sub

Private Sub ExeSaveData()
'���ܣ���������
    Dim lng�ļ�ID As Long, strContent As String
    Dim i As Long, strSQL As String
    Dim blnAllNull As Boolean
    
    On Error GoTo errH
    lng�ļ�ID = cboDate.ItemData(cboDate.ListIndex)
    
    With vsItem
        For i = .FixedRows To .Rows - 1
            '�к�|���|��ע||...,ĩβ��||,�����עΪ����Ҫ��һ���ո�
            strContent = strContent & .RowData(i) & "|" & IIf(Trim(.TextMatrix(i, col���)) = "", " ", Trim(.TextMatrix(i, col���))) & "|" & IIf(Trim(.TextMatrix(i, col��ע)) = "", " ", Trim(.TextMatrix(i, col��ע))) & "||"
            
        Next
        If blnAllNull Then strContent = ""
    End With
    
    strSQL = "Zl_·�������ļ�_Update(" & lng�ļ�ID & ",'" & strContent & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    stbThis.Panels(2).Text = "���ݱ���ɹ���"
    Call SetVsitemEdit(False)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ExeDelete()
'���ܣ�ִ������ɾ������
    Dim lngId As Long, strSQL As String
        
    If MsgBox("��ȷ��Ҫɾ��" & cboDate.Text & "�ı�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    On Error GoTo errH
    
    lngId = cboDate.ItemData(cboDate.ListIndex)
    
    strSQL = "Zl_·�������ļ�_Delete(" & lngId & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    stbThis.Panels(2).Text = cboDate.Text & "�ı���ɾ���ɹ���"
    
    Call FillDate
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ExeModify()
'���ܣ�ִ���޸����ݲ���
    Call SetVsitemEdit(True)

End Sub

Private Sub ExeNew()
'���ܣ���������
    
    frmReport1Add.mlng·��ID = mlng·��ID
    frmReport1Add.Show vbModal, gfrmMain
    
    If frmReport1Add.mblnOK Then
        Call SetVsitemEdit(True)
        Call FillDate(frmReport1Add.mstr�ڼ�)
    End If
End Sub

Private Sub FillDate(Optional ByVal str�ڼ� As String)
'���ܣ����ص�ǰ·������ڼ�
'      str�ڼ�:ȱʡ�ĵ�ǰ�ڼ�
    Dim strSQL As String
    Dim i As Long
 
    strSQL = "Select �ڼ�,ID,��ʼʱ��,����ʱ�� From ·�������ļ� Where ����ID = 1 And ·��ID = [1] Order by �ڼ� Desc"
    On Error GoTo errH
    Set mrsDate = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng·��ID)
        
    cboDate.Clear
    For i = 1 To mrsDate.RecordCount
        cboDate.AddItem Mid(mrsDate!�ڼ�, 1, 4) & "��" & Mid(mrsDate!�ڼ�, 5, 2) & "��"
        cboDate.ItemData(cboDate.NewIndex) = mrsDate!ID
        mrsDate.MoveNext
    Next
    If cboDate.ListCount > 0 Then
        If str�ڼ� = "" Then
            cboDate.ListIndex = 0
        Else
            cbo.Locate cboDate, Mid(str�ڼ�, 1, 4) & "��" & Mid(str�ڼ�, 5, 2) & "��"
        End If
    Else
        dkpMain.Panes(1).Title = "���֣�" & mstr·������
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub FillStructure()
'���ܣ�ˢ�±���ṹ
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long
 
    strSQL = "Select �к�, ��Ŀ���, ��Ŀ�ı�1, ��Ŀ�ı�2, Sql�ı� From ·������ṹ Where ����id = 1 Order By �к�"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        With vsItem
            .Rows = .FixedRows
            .Rows = .FixedRows + rsTmp.RecordCount
            
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val("" & rsTmp!�к�)
                .TextMatrix(i, COL���) = "" & rsTmp!��Ŀ���
                .TextMatrix(i, COL��Ŀ�ı�1) = "" & rsTmp!��Ŀ�ı�1
                .TextMatrix(i, COL��Ŀ�ı�2) = NVL(rsTmp!��Ŀ�ı�2, rsTmp!��Ŀ�ı�1)
                .Cell(flexcpData, i, col���) = "" & rsTmp!SQL�ı�
                .MergeRow(i) = True
                
                If Not IsNull(rsTmp!SQL�ı�) Then
                    .Cell(flexcpBackColor, i, col���) = &HFFEFDF
                End If
                
                rsTmp.MoveNext
            Next
            
            .MergeCol(COL���) = True
            .MergeCol(COL��Ŀ�ı�1) = True
            .MergeCol(col���) = False
            .MergeCol(col��ע) = False
        End With
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function zlRefresh() As Long
'���ܣ�ˢ�±�������
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long, lngId As Long
    
    If cboDate.ListIndex < 0 Then
        With vsItem
            .Redraw = flexRDNone
            For i = 1 To .Rows - 1
                .TextMatrix(i, col���) = ""
                .TextMatrix(i, col��ע) = ""
            Next
            .Redraw = flexRDDirect
        End With
    Else
        lngId = cboDate.ItemData(cboDate.ListIndex)
        
        strSQL = "Select a.�к�, a.��Ŀֵ, a.��ע" & vbNewLine & _
                "From ·�������¼ A" & vbNewLine & _
                "Where a.�ļ�id = [1]" & vbNewLine & _
                "Order By �к�"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngId)
        If rsTmp.RecordCount > 0 Then
            With vsItem
                .Redraw = flexRDNone
                For i = 1 To .Rows - 1
                    rsTmp.Filter = "�к�=" & .RowData(i)
                    If rsTmp.RecordCount > 0 Then
                        .TextMatrix(i, col���) = "" & rsTmp!��Ŀֵ
                        .TextMatrix(i, col��ע) = "" & rsTmp!��ע
                    Else
                        .TextMatrix(i, col���) = ""
                        .TextMatrix(i, col��ע) = ""
                    End If
                Next
                .Redraw = flexRDDirect
            End With
        End If
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub vsItem_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = col��� Or Col = col��ע) Then
        Cancel = True
    End If
End Sub

Private Sub ExeDefineSQL()
    Dim lng�к� As Long, strSQL�ı� As String
    
    lng�к� = Val(vsItem.RowData(vsItem.Row))
    
    If lng�к� <> 0 Then
        '1,4,17���������У����ܶ���SQL
        If lng�к� = 1 Or lng�к� = 4 Or lng�к� = 17 Then
            MsgBox "��ǰ���Ǳ����У����ܶ���SQL��", vbInformation, gstrSysName
            Exit Sub
        End If
    
        strSQL�ı� = vsItem.Cell(flexcpData, vsItem.Row, col���)
        Call frmReportSQLSet.ShowMe(gfrmMain, lng�к�, strSQL�ı�)
        vsItem.Cell(flexcpData, vsItem.Row, col���) = strSQL�ı�
    End If
End Sub

Private Sub vsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        With vsItem
            If (.Col = col��� Or .Col = col��ע) And mblnEdit Then .TextMatrix(.Row, .Col) = ""
        End With
    End If
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call GoNextCell
    End If
End Sub

Private Sub GoNextCell()
    If vsItem.Row < vsItem.Rows - 1 Then
        vsItem.Row = vsItem.Row + 1
        If vsItem.Col = vsItem.Cols - 1 Then vsItem.Col = col���
    End If
End Sub

Private Sub vsItem_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call GoNextCell
    End If
End Sub

Private Sub ExeGetData(Optional ByVal lngRow As Long)
'���ܣ�����Ԥ����SQL��ȡ������䵥Ԫ��
    Dim strSQL As String, strPati As String, strPatiPage As String
    Dim rsTmp As ADODB.Recordset
    Dim DatBegin As Date, DatEnd As Date, i As Long, lngS As Long, lngE As Long, lngErrCnt As Long
    Dim lng��סԺ���� As Long, lng�ܷ��� As Long, lng���˴� As Long, lng������ As Long
    Dim blnGeted�ܷ��� As Boolean   '�Ƿ�ȡ���ܷ���
    Dim varPati As Variant, varPage As Variant
    Dim strParTable As String, strTablePati As String, strTablePage As String
    Dim strTempSQL As String
    Dim intMaxPati As Integer, intMaxPage As Integer

    On Error GoTo errH
    '��ȡ������Ϣ
    If mrsPati Is Nothing Or mstrǰһ�ڼ� <> mstr��ǰ�ڼ� Then
        DatBegin = mrsDate!��ʼʱ��
        DatEnd = mrsDate!����ʱ��
        strSQL = "Select a.����id, a.��ҳid, a.����id" & vbNewLine & _
                "From (Select Row_Number() Over(Partition By a.����id Order By a.��ҳid Desc, Decode(a.��¼��Դ,4,1,a.��¼��Դ) Desc,Sign(a.�������-10), Decode(a.�������,3,0,13,10,a.�������) Desc) As Top, a.����id, a.��ҳid," & vbNewLine & _
                "              a.����id" & vbNewLine & _
                "       From ������ϼ�¼ A, ������ҳ B" & vbNewLine & _
                "       Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.��ϴ��� = 1 And a.������� In (1, 2, 3, 11, 12, 13) And" & vbNewLine & _
                "             b.��Ժ���� Between [2] And [3] And a.����id is Not Null) A, �ٴ�·������ B" & vbNewLine & _
                "Where Top = 1 And a.����id = b.����id And b.����=0 And b.·��ID = [1]"
        Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng·��ID, DatBegin, DatEnd)
    End If
    
    If mrsPati.RecordCount > 0 Then
        mrsPati.MoveFirst
        For i = 1 To mrsPati.RecordCount
            strPatiPage = strPatiPage & "," & mrsPati!����ID & ":" & mrsPati!��ҳID
            If InStr(strPati & ",", "," & mrsPati!����ID & ",") = 0 Then
                strPati = strPati & "," & mrsPati!����ID
            End If
            mrsPati.MoveNext
        Next
        strPatiPage = Mid(strPatiPage, 2)
        strPati = Mid(strPati, 2)
        
        If Len(strPatiPage) > 4000 Then
            varPage = FuncGetTable(strPatiPage, 1, strTablePage, intMaxPage)
        End If
        If Len(strPati) > 4000 Then
            varPati = FuncGetTable(strPati, 0, strTablePati, intMaxPati)
        End If
    End If
    If strPatiPage = "" Then
        MsgBox "��ǰ�ڼ�û�з��ϲ���[" & mstr·������ & "]�����Ĳ��ˡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    lng���˴� = UBound(Split(strPatiPage, ",")) + 1
    lng������ = UBound(Split(strPati, ",")) + 1
    
    If lngRow = 0 Then
        lngS = vsItem.FixedRows
        lngE = vsItem.Rows - 1
    Else
        lngS = lngRow
        lngE = lngRow
    End If
    
    On Error Resume Next '���������¼���¼�������
    With vsItem
        For i = lngS To lngE
            stbThis.Panels(2).Text = "���ڶ�ȡ��" & i & "������"
            Me.Refresh
            .Cell(flexcpData, i, col��ע) = ""  '���ǰһ�ε�
            .Cell(flexcpBackColor, i, col��ע) = vbWhite    'ȡ������󣬱�ע�е�ɫΪǳ��ɫ
                    
            strSQL = .Cell(flexcpData, i, col���)
            If strSQL <> "" Then
                .TextMatrix(i, col���) = ""
                '1:һ��Ч��ָ��
                
                Select Case .RowData(i) '�����к�
                Case 2 '����ƽ��סԺ����,��סԺ�������ڼ���"�վ�����"
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, mlng·��ID)
                    Else
                        '�����ŵ�������δ�����Ĳ���������ǰ��
                        Call FuncReDoSQLNum(strSQL, 2, 2)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '��������ƶ�һλ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then
                            .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                            If rsTmp.Fields.count > 1 Then lng��סԺ���� = Val("" & rsTmp.Fields(1).Value)
                        End If
                    End If
                Case 3  '��ǰƽ��סԺ��
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 2)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '��������ƶ�һλ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                    End If
                    
                '4:����Ч��ָ��
                Case 5  '������,������,��ת��
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng���˴�, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '��������ƶ�2λ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng���˴�, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    .TextMatrix(i + 1, col���) = "": .TextMatrix(i + 2, col���) = ""
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then
                            .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                            If rsTmp.Fields.count > 1 Then .TextMatrix(i + 1, col���) = "" & rsTmp.Fields(1).Value
                            If rsTmp.Fields.count > 2 Then .TextMatrix(i + 2, col���) = "" & rsTmp.Fields(2).Value
                        End If
                    End If
                Case 6, 7   '������,��ת��(��������ˣ�ȱʡδ����)
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 2)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '��������ƶ�һλ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                    End If
                Case 8  'Ժ�ڸ�Ⱦ��
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng���˴�, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '��������ƶ�2λ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng���˴�, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                    End If
                Case 9  '������λ��Ⱦ��(��������ˣ�ȱʡδ����)
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng���˴�, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '��������ƶ�2λ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng���˴�, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                    End If
                Case 10 '14����סԺ��,31����סԺ��
                    .TextMatrix(i + 1, col���) = ""
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPati, strPatiPage, lng���˴�, mlng·��ID)
                    ElseIf Len(strPati) <= 4000 Then
                        Call FuncReDoSQLNum(strSQL, 3, 4)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 3) '��������ƶ�3λ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([4]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPati, lng���˴�, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    Else
                        Call FuncReDoSQLNum(strSQL, 3, 4)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 3) '��������ƶ�3λ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([4]))"), strTempSQL)
                        
                        Call FuncReDoSQLNum(strSQL, 2, 13) '����page���ȡ��10������Pati����λ��
                        strTempSQL = strTablePati
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPati, 12)  '��������ƶ�12λ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list([13]))"), strTempSQL)
                        
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng���˴�, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)), CStr(varPati(0)), CStr(varPati(1)), _
                            CStr(varPati(2)), CStr(varPati(3)), CStr(varPati(4)), CStr(varPati(5)), CStr(varPati(6)), CStr(varPati(7)), CStr(varPati(8)), CStr(varPati(9)))
    
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                        If rsTmp.Fields.count > 1 Then .TextMatrix(i + 1, col���) = "" & rsTmp.Fields(1).Value
                    End If
                Case 11, 12 '31����סԺ��,�Ǽƻ��ط������ҷ�����(��������ˣ�ȱʡδ����)
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng���˴�, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '��������ƶ�2λ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng���˴�, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                    End If
               
                    
                '-----------------------------------------------
                Case 13 '����֢�����ʣ���ǰ��λ����
                    .TextMatrix(i + 1, col���) = "": .TextMatrix(i + 2, col���) = "": .TextMatrix(i + 3, col���) = ""
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng���˴�, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '��������ƶ�2λ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng���˴�, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then
                            .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                            If rsTmp.Fields.count > 1 Then .TextMatrix(i + 1, col���) = "" & rsTmp.Fields(1).Value
                            If rsTmp.Fields.count > 2 Then .TextMatrix(i + 2, col���) = "" & rsTmp.Fields(2).Value
                            If rsTmp.Fields.count > 3 Then .TextMatrix(i + 3, col���) = "" & rsTmp.Fields(3).Value
                        End If
                    End If
                Case 14, 15, 16 '(��������ˣ�ȱʡδ����)
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 2)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '��������ƶ�һλ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                    End If

                '-----------------------------------------------------
                '17:����������ָ��
                Case 18 'סԺ����������(��������ˣ�ȱʡδ����)
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, DatBegin, DatEnd, mlng·��ID)
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                    End If
                
                Case 19 '����·���Ļ������˴���
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 2)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '��������ƶ�һλ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                    End If
                
                Case 20 '���·�����˴���
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 2)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '��������ƶ�һλ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                    End If
                    
                Case 21 '���첡����
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 2)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '��������ƶ�һλ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                    End If

                '------------------------------------------------
                Case 23 'ʹ����������ҩ��Ļ��߱���
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng���˴�, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '��������ƶ�2λ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng���˴�, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                    End If
                
                Case 24 'ʹ�ÿ����ص�ƽ������
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 2)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '��������ƶ�һλ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                    End If
                
                Case 25 '�����ִξ�����,�ܷ���(��2��ȡ��סԺ����) '���ֻ��һ����Ԫ����lng��סԺ����û��ֵ����ҪSQL���Լ�ȡ
                    blnGeted�ܷ��� = True
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng���˴�, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '��������ƶ�2λ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng���˴�, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                        If rsTmp.Fields.count > 1 Then lng�ܷ��� = Val("" & rsTmp.Fields(1).Value)
                    End If
                
                Case 26 '�������վ�����(��������ˣ�ȱʡδ����)
                    If lng��סԺ���� = 0 Then lng��סԺ���� = Get����(strPatiPage, lng���˴�, 2, strTablePage, intMaxPage, varPage)
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng��סԺ����, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '��������ƶ�2λ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��סԺ����, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                    End If
                
                Case 27 '����ҩ����ñ�
                    If blnGeted�ܷ��� = False Then lng�ܷ��� = Get����(strPatiPage, lng���˴�, 25, strTablePage, intMaxPage, varPage): blnGeted�ܷ��� = True
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng�ܷ���, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '��������ƶ�2λ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ܷ���, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                    End If
                
                Case 28 '�Ĳķ��ñ�
                    If blnGeted�ܷ��� = False Then lng�ܷ��� = Get����(strPatiPage, lng���˴�, 25, strTablePage, intMaxPage, varPage): blnGeted�ܷ��� = True
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng�ܷ���, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '��������ƶ�2λ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ܷ���, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                    End If
                
                Case 29 '�����ñ�
                    If blnGeted�ܷ��� = False Then lng�ܷ��� = Get����(strPatiPage, lng���˴�, 25, strTablePage, intMaxPage, varPage): blnGeted�ܷ��� = True
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng�ܷ���, mlng·��ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '��������ƶ�2λ
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ܷ���, mlng·��ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col���) = "" & rsTmp.Fields(0).Value
                    End If
                
                End Select
            
                If gcnOracle.Errors.count <> 0 Then
                    .Cell(flexcpData, i, col��ע) = CStr(gcnOracle.Errors(0).Description)
                    .Cell(flexcpBackColor, i, col��ע) = &HC0FFFF
                    gcnOracle.Errors.Clear
                    stbThis.Panels(2).Text = "��" & i & "�ж�ȡ���ݳ���" & CStr(gcnOracle.Errors(0).Description)
                    lngErrCnt = lngErrCnt + 1
                    Me.Refresh
                End If
                
                'û�ж���SQL�ĵ�Ԫ��ȡ��
            Else
                Select Case .RowData(i)
                    Case 18 'סԺ����������
                        .TextMatrix(i, col���) = lng������
                    Case 26 '�������վ�����
                        If blnGeted�ܷ��� = False Then lng�ܷ��� = Get����(strPatiPage, lng���˴�, 25, strTablePage, intMaxPage, varPage): blnGeted�ܷ��� = True
                        If lng��סԺ���� = 0 Then lng��סԺ���� = Get����(strPatiPage, lng���˴�, 2, strTablePage, intMaxPage, varPage)
                        .TextMatrix(i, col���) = Format(lng�ܷ��� / lng��סԺ����, "#######0.00")
                End Select
            End If
        Next
    End With
    
    If lngErrCnt <> 0 Then
        stbThis.Panels(2).Text = "��" & lngErrCnt & "�����ݶ�ȡ����,����Ƶ���ɫ��Ԫ��ɲ鿴��ϸ����"
    Else
        stbThis.Panels(2).Text = "��������ݶ�ȡ������"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function Get����(ByVal strPatiPage As String, ByVal lng���˴� As Long, ByVal lngRow As Long, ByVal strTablePage As String, _
    ByVal intMaxPage As Integer, ByVal varPage As Variant) As Long
'���ܣ���ȡ�ܷ���(��25��)����סԺ����(��2��)
'������strPatiPage:���˼���ҳID�б�:1121:1,1122:1
'      lngRow:�����кţ���ת��Ϊ����к�
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strTempSQL As String

    On Error Resume Next
    With vsItem
        For i = .FixedRows To .Rows - 1
            If .RowData(i) = lngRow Then Exit For
        Next
    
        strSQL = .Cell(flexcpData, i, col���)
        If strSQL <> "" Then
            If Len(strPatiPage) <= 4000 Then
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng���˴�)
            Else
                '�����ŵ�������δ�����Ĳ���������ǰ��
                Call FuncReDoSQLNum(strSQL, 2, 2)
                strTempSQL = strTablePage
                Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '��������ƶ�һλ
                strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng���˴�, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                    CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
            End If
            If gcnOracle.Errors.count = 0 Then
                If rsTmp.RecordCount > 0 Then
                    If lngRow = 25 Or lngRow = 2 Then
                        If rsTmp.Fields.count > 1 Then Get���� = Val("" & rsTmp.Fields(1).Value)  '��ȡ�ܷ��ã���סԺ����,�̶�ȡ��2���ֶ�
                    Else
                        Get���� = Val("" & rsTmp.Fields(0).Value)
                    End If
                End If
            Else
                .Cell(flexcpData, i, col��ע) = CStr(gcnOracle.Errors(0).Description)
                .Cell(flexcpBackColor, i, col��ע) = &HC0FFFF
                gcnOracle.Errors.Clear
            End If
        End If
    End With
End Function

Private Sub vsItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsItem
        If .MouseCol >= .FixedCols And .MouseRow >= .FixedRows Then
            Dim strErr As String
            If .Col = col��ע Then
                strErr = .Cell(flexcpData, .MouseRow, col��ע)
                If strErr <> "" Then
                    Call zlCommFun.ShowTipInfo(.Hwnd, strErr)
                Else
                    Call zlCommFun.ShowTipInfo(.Hwnd, "")
                End If
            Else
                Call zlCommFun.ShowTipInfo(.Hwnd, "")
            End If
        End If
    End With
End Sub

Private Function FuncGetTable(ByVal strPara As String, ByVal bytFunc As Byte, ByRef strTableOut As String, ByRef intMaxIdx As Integer) As Variant
'���ܣ����ڶ�̬�ڴ��İ󶨲��������Ĵ���
'������strPar ������ ����4000,��ַָ���Ĭ����","
'   bytFunc=0 �ڴ��f_Num2list;bytFunc=1 :f_Num2list2
'���أ�һ���ַ������飬���10��Ԫ��
'    strTableOut=�����붯̬�ڴ���Ч��SQL���(UNION ALL ����)
'   intMaxIdx = ���ز��֮��õ����������
    Dim varPara As Variant
    Dim strParTable As String
    
    varPara = Array()
    
    If bytFunc = 0 Then
        strParTable = "Select Column_Value From Table(f_Num2list([1]))"
    Else
        strParTable = "Select C1, C2 From Table(f_Num2list2([1]))"
    End If
    varPara = GetParTable(strPara, strParTable, strTableOut, intMaxIdx)
    strTableOut = "(" & strTableOut & ")"
                     
    FuncGetTable = varPara
End Function


Private Sub FuncMoveSQLNum(ByRef strSQL As String, ByVal intBegin As Integer, ByVal intEnd As Integer, ByVal intMoveLen As Integer)
'����:��SQL�в��������������ƶ�����ǰ�ƶ�
'����:intBegin,intEnd=�����Ĳ�����ŵı�����[intBegin,intEnd]
'intMoveLen=ƫ���� >0����ƶ�,<0��ǰ�ƶ�
    Dim i As Integer
    If intMoveLen > 0 Then
    i = intEnd
        Do While i >= intBegin
            strSQL = Replace(strSQL, "[" & i & "]", "[" & (i + intMoveLen) & "]")
            i = i - 1
        Loop
    Else
        For i = intBegin To intEnd
             strSQL = Replace(strSQL, "[" & i & "]", "[" & (i + intMoveLen) & "]")
        Next
    End If
End Sub

Private Sub FuncReDoSQLNum(ByRef strSQL As String, ByVal intBegin As Integer, ByVal intEnd As Integer)
'����:��SQL�в���ֵ���ȳ���4000�Ĳ�����ŵ�������SQL�����Ĳ������
'����:
'intBegin:��Ҫ����������ʼλ��
'intEnd:��Ҫ��������ĩβλ�ã�intEnd>=intBegin��
'intBegin-1:�ò���ֵ��Ӧ���ȳ���4000 ;��intEnd + 1��������ʱ�洢
'����:�������SQL
    strSQL = Replace(strSQL, "[" & (intBegin - 1) & "]", "[" & (intEnd + 1) & "]")
    Call FuncMoveSQLNum(strSQL, intBegin, (intEnd + 1), -1)
End Sub
