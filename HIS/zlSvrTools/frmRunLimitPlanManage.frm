VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Begin VB.Form frmRunLimitPlanManage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12675
   Icon            =   "frmRunLimitPlanManage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.TreeView tvwPlanTree 
      Height          =   5235
      Left            =   45
      TabIndex        =   0
      Top             =   975
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   9234
      _Version        =   393217
      Indentation     =   706
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img24"
      Appearance      =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPlanDetail 
      Height          =   5115
      Left            =   3825
      TabIndex        =   1
      Top             =   975
      Width           =   8235
      _cx             =   14526
      _cy             =   9022
      Appearance      =   1
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
      BackColor       =   16774866
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16774866
      BackColorAlternate=   16774866
      GridColor       =   -2147483633
      GridColorFixed  =   15984570
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   8
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmRunLimitPlanManage.frx":6852
      ScrollTrack     =   0   'False
      ScrollBars      =   1
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
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
   Begin MSComctlLib.ImageList img24 
      Left            =   2970
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitPlanManage.frx":68E6
            Key             =   "enabled"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitPlanManage.frx":79F8
            Key             =   "enabledLock"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitPlanManage.frx":947A
            Key             =   "disabled"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitPlanManage.frx":AEFC
            Key             =   "disabledLock"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgPlanDetail 
      Left            =   3075
      Top             =   1500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   97
      ImageHeight     =   30
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitPlanManage.frx":C97E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager ImageMan 
      Left            =   3765
      Top             =   165
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmRunLimitPlanManage.frx":10C54
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   3165
      Top             =   135
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Visible         =   0   'False
      Begin VB.Menu mnuViewShow 
         Caption         =   "��ʾͣ�÷���(&S)"
      End
   End
   Begin VB.Menu mnuPlanName 
      Caption         =   "��������"
      Visible         =   0   'False
      Begin VB.Menu mnuPlanNameNew 
         Caption         =   "��������(&N)"
      End
      Begin VB.Menu mnuPlanNameUpdate 
         Caption         =   "�޸ķ���(&U)"
      End
      Begin VB.Menu mnuPlanNameRemove 
         Caption         =   "ɾ������(&R)"
      End
      Begin VB.Menu mnuPlanNameSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlanNameStart 
         Caption         =   "���÷���(&S)"
      End
      Begin VB.Menu mnuPlanNameStop 
         Caption         =   "ͣ�÷���(&T)"
      End
   End
   Begin VB.Menu mnuPlanDetail 
      Caption         =   "��������"
      Visible         =   0   'False
      Begin VB.Menu mnuPlanDetailAdd 
         Caption         =   "����ʱ���(&A)"
      End
      Begin VB.Menu mnuPlanDetailModify 
         Caption         =   "�޸�ʱ���(&M)"
      End
      Begin VB.Menu mnuPlanDetailDel 
         Caption         =   "ɾ��ʱ���(&D)"
      End
   End
End
Attribute VB_Name = "frmRunLimitPlanManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsPlan As ADODB.Recordset
Private mlngPlanNo As Long
Private mobjBar As CommandBar
Private mobjMenu As CommandBarPopup
Private mobjPopup As CommandBarPopup
Private mobjControl As CommandBarControl
Private Const vsfTitleBackColor = &HF0E5BD  '�������ݱ����ⱳ����ɫ
Private Const vsfContentBackColor = &HFFFAE4 '�������ݱ�����ݲ���ǳɫ����ɫ
Private Const HighlightForeColor = &H80000005  '����ǰ��ɫ
Private Const HighlightBackColor = &H8000000D  '��������ɫ
Private Const vsfTitleHeight = 500
Private Const vsfRowHeight = 1000
Private Enum PlanDetailTitle
    PDT_���� = 0
    PDT_ʱ���1 = 1
    PDT_ʱ�����չ = 2
End Enum
Private Enum PlanDetail
    PD_���� = 0
    PD_������ = 1
    PD_����һ = 2
    PD_���ڶ� = 3
    PD_������ = 4
    PD_������ = 5
    PD_������ = 6
    PD_������ = 7
End Enum

Private Enum CbsMainId
    CMI_Exit = 11
    CMI_NewPlan = 21
    CMI_UpdatePlan = 22
    CMI_RemovePlan = 23
    CMI_StartPlan = 24
    CMI_StopPlan = 25
    CMI_AddTime = 26
    CMI_EditTime = 27
    CMI_DeleteTime = 28
    CMI_ShowStopPlan = 31
End Enum

Public Sub ShowMe(Optional ByVal lngPlanNo As Long)
    '�����lngPlanNo�Ļ�����ѡ�ж�Ӧ����
    mlngPlanNo = lngPlanNo
    If mlngPlanNo = 0 Then mlngPlanNo = 1
    Me.Show vbModal, frmMDIMain
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errH
    Select Case Control.id
        Case CMI_Exit
            '�˳�
            Unload Me
        Case CMI_NewPlan
            '��������
            Call NewPlan
        Case CMI_UpdatePlan
            '�޸ķ���
            Call UpdatePlan
        Case CMI_RemovePlan
            'ɾ������
            Call RemovePlan
        Case CMI_StartPlan
            '���÷���
            Call StartPlan
        Case CMI_StopPlan
            'ͣ�÷���
            Call StopPlan
        Case CMI_AddTime
            '����ʱ���
            Call AddTime
        Case CMI_EditTime
            '�޸�ʱ���
            Call EditTime
        Case CMI_DeleteTime
            'ɾ��ʱ���
            Call DeleteTime
        Case CMI_ShowStopPlan
            '��ʾͣ�÷���
            Control.Checked = Not Control.Checked
            Call ShowStopPlan
    End Select
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case CMI_UpdatePlan
            '�޸ķ���
            Control.Enabled = mnuPlanNameUpdate.Enabled
        Case CMI_RemovePlan
            'ɾ������
            Control.Enabled = mnuPlanNameRemove.Enabled
        Case CMI_StartPlan
            '���÷���
            Control.Enabled = mnuPlanNameStart.Enabled
        Case CMI_StopPlan
            'ͣ�÷���
            Control.Enabled = mnuPlanNameStop.Enabled
        Case CMI_AddTime
            '����ʱ���
            Control.Enabled = mnuPlanDetailAdd.Enabled
        Case CMI_EditTime
            '�޸�ʱ���
            Control.Enabled = mnuPlanDetailModify.Enabled
        Case CMI_DeleteTime
            'ɾ��ʱ���
            Control.Enabled = mnuPlanDetailDel.Enabled
    End Select
End Sub

Private Sub Form_Load()
    Call InitCbsMain
    Call FillPlanList
    Call FormatVsfPlan
End Sub

Private Sub InitCbsMain()
    With CommandBarsGlobalSettings
        Set .App = App
        .ResourceFile = .OcxPath & "\XTPResourceZhCn.dll"                                       '��������������Դ�ļ�
        .ColorManager.SystemTheme = xtpSystemThemeAuto                                          '�ؼ��������ɫ����
    End With
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False                                                         '�����ڹ������Ҳ���ʾѡ�ť,��ʹ�������㹻��
        .ToolBarAccelTips = True                                                                '��ʾ��ť��ʾ
        .AlwaysShowFullMenus = False                                                            '�����õĲ˵���������
        .UseFadedIcons = True                                                                   'ͼ����ʾΪ��ɫЧ��
        .IconsWithShadow = True                                                                 '���ָ�������ͼ����ʾ��ӰЧ��
        .UseDisabledIcons = True                                                                '��������ť����ʱͼ����ʾΪ������ʽ
        .LargeIcons = True                                                                      '��������ʾΪ��ͼ��
        .SetIconSize True, 24, 24                                                               '���ô�ͼ��ĳߴ�
        .SetIconSize False, 16, 16                                                              '����Сͼ��ĳߴ�
    End With
    
    With cbsMain
        .VisualTheme = xtpThemeOffice2003                                                       '���ÿؼ���ʾ���
        .EnableCustomization False                                                              '�Ƿ������Զ�������
        Set cbsMain.Icons = ImageMan.Icons                                                      '���ù�����ͼ��ؼ�
    End With                                                                                    '�˵�����Զ������ҿ�Ȳ���ʱҲ������
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set mobjMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 1, "�ļ�(&F)", -1, False) '�˵�������
    With mobjMenu.CommandBar.Controls
        Set mobjControl = .Add(xtpControlButton, 11, "�˳�(&X)")
    End With

    Set mobjMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 2, "�༭(&E)", -1, False)
    With mobjMenu.CommandBar.Controls
        Set mobjControl = .Add(xtpControlButton, 21, "��������(&N)")
        Set mobjControl = .Add(xtpControlButton, 22, "�޸ķ���(&U)")
        Set mobjControl = .Add(xtpControlButton, 23, "ɾ������(&R)")
        
        Set mobjControl = .Add(xtpControlButton, 24, "���÷���(&S)")
        mobjControl.BeginGroup = True
        Set mobjControl = .Add(xtpControlButton, 25, "ͣ�÷���(&T)")
        
        Set mobjControl = .Add(xtpControlButton, 26, "����ʱ���(&A)")
        mobjControl.BeginGroup = True
        Set mobjControl = .Add(xtpControlButton, 27, "�޸�ʱ���(&M)")
        Set mobjControl = .Add(xtpControlButton, 28, "ɾ��ʱ���(&D)")
    End With
    
    Set mobjMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 3, "�鿴(&V)", -1, False)
    With mobjMenu.CommandBar.Controls
        Set mobjControl = .Add(xtpControlButton, 31, "��ʾͣ�÷���(&S)")
        mobjControl.Checked = False
    End With
'
    Set mobjBar = cbsMain.Add("������", xtpBarTop)                                               '����������
    With mobjBar.Controls
        Set mobjControl = .Add(xtpControlButton, 21, "��������")
        mobjControl.Style = xtpButtonIconAndCaption
        Set mobjControl = .Add(xtpControlButton, 22, "�޸ķ���")
        mobjControl.Style = xtpButtonIconAndCaption
        Set mobjControl = .Add(xtpControlButton, 23, "ɾ������")
        mobjControl.Style = xtpButtonIconAndCaption
        
        Set mobjControl = .Add(xtpControlButton, 24, "���÷���")
        mobjControl.Style = xtpButtonIconAndCaption
        mobjControl.BeginGroup = True
        Set mobjControl = .Add(xtpControlButton, 25, "ͣ�÷���")
        mobjControl.Style = xtpButtonIconAndCaption
        
        Set mobjControl = .Add(xtpControlButton, 26, "����ʱ���")
        mobjControl.Style = xtpButtonIconAndCaption
        mobjControl.BeginGroup = True
        Set mobjControl = .Add(xtpControlButton, 27, "�޸�ʱ���")
        mobjControl.Style = xtpButtonIconAndCaption
        Set mobjControl = .Add(xtpControlButton, 28, "ɾ��ʱ���")
        mobjControl.Style = xtpButtonIconAndCaption
        
        Set mobjControl = .Add(xtpControlButton, 11, "�˳�")
        mobjControl.Style = xtpButtonIconAndCaption
        mobjControl.BeginGroup = True
    End With
End Sub

Private Sub FillPlanList()
    '������·������б�
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim objNode As Node
    Dim i As Long
    
    On Error GoTo errH
    strSql = "Select ���, ����, �Ƿ�����, ���� From ZlRunLimit Order by �Ƿ����� Desc, ���"
    Set mrsPlan = gclsBase.OpenSQLRecord(gcnOracle, strSql, Me.Caption)
    tvwPlanTree.Nodes.Clear
    With mrsPlan
        Do While Not .EOF
            If !�Ƿ����� = 1 Then
                If !���� = "Ԥ�跽��" Then
                    Set objNode = tvwPlanTree.Nodes.Add(, , "K_" & !���, !����, "enabledLock")
                Else
                    Set objNode = tvwPlanTree.Nodes.Add(, , "K_" & !���, !����, "enabled")
                End If
                objNode.Tag = !���� & ""
            ElseIf mnuViewShow.Checked Then
                If !���� = "Ԥ�跽��" Then
                    Set objNode = tvwPlanTree.Nodes.Add(, , "K_" & !���, !����, "disabledLock")
                Else
                    Set objNode = tvwPlanTree.Nodes.Add(, , "K_" & !���, !����, "disabled")
                End If
                objNode.Tag = !���� & ""
            End If
            .MoveNext
        Loop
        On Error Resume Next
            tvwPlanTree.Nodes("K_" & mlngPlanNo).Selected = True
            Call tvwPlanTree_NodeClick(tvwPlanTree.Nodes("K_" & mlngPlanNo))
        If err.Number <> 0 Then
            If tvwPlanTree.Nodes.Count > 0 Then
                mlngPlanNo = Split(tvwPlanTree.Nodes(1).Key, "_")(1)
                tvwPlanTree.Nodes("K_" & mlngPlanNo).Selected = True
                Call tvwPlanTree_NodeClick(tvwPlanTree.Nodes("K_" & mlngPlanNo))
            Else
                Call ClearPlanDetail
                Call SetEnabled(False)
            End If
            err.Clear
        End If
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub FormatVsfPlan()
    '�����Ҳ෽��չʾ����ʽ
    With vsfPlanDetail
        .Cell(flexcpPicture, 0, 0) = imgPlanDetail.ListImages(1).Picture
        .GridLines = flexGridNone
        .rowHeight(PD_����) = vsfTitleHeight
        .rowHeight(PD_������) = vsfRowHeight
        .rowHeight(PD_����һ) = vsfRowHeight
        .rowHeight(PD_���ڶ�) = vsfRowHeight
        .rowHeight(PD_������) = vsfRowHeight
        .rowHeight(PD_������) = vsfRowHeight
        .rowHeight(PD_������) = vsfRowHeight
        .rowHeight(PD_������) = vsfRowHeight
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    tvwPlanTree.Height = Me.ScaleHeight - tvwPlanTree.Top - 50
    vsfPlanDetail.Height = Me.ScaleHeight - vsfPlanDetail.Top - 50
    vsfPlanDetail.Left = tvwPlanTree.Left + tvwPlanTree.Width + 50
    vsfPlanDetail.Width = Me.ScaleWidth - vsfPlanDetail.Left - 50
    Call AdjustFormDisplay
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub AdjustFormDisplay()
    With vsfPlanDetail
        .Select 0, 0, .Rows - 1, .Cols - 1
        .CellBorder &HE9D2A5, 1, 0, 1, 2, 2, 2
        .Cell(flexcpBackColor, PD_����, PDT_����, 0, .Cols - 1) = vsfTitleBackColor
        .Cell(flexcpBackColor, PD_����, PDT_����, .Rows - 1, 0) = vsfTitleBackColor
        .Cell(flexcpBackColor, PD_������, PDT_ʱ���1, PD_������, .Cols - 1) = vsfContentBackColor
        .Cell(flexcpBackColor, PD_���ڶ�, PDT_ʱ���1, PD_���ڶ�, .Cols - 1) = vsfContentBackColor
        .Cell(flexcpBackColor, PD_������, PDT_ʱ���1, PD_������, .Cols - 1) = vsfContentBackColor
        .Cell(flexcpBackColor, PD_������, PDT_ʱ���1, PD_������, .Cols - 1) = vsfContentBackColor
    End With
End Sub

Private Sub AddTime()
'����ʱ���
    Dim strTimeStart As String, strTimeStop As String
    Dim lngRow As Long, lngCol As Long
    Dim j As Long
    
    If tvwPlanTree.SelectedItem Is Nothing Then Exit Sub
    If tvwPlanTree.SelectedItem.Text = "Ԥ�跽��" Then Exit Sub
    With vsfPlanDetail
        If .Row = 0 Or .Row = 8 Then Exit Sub
        lngRow = .Row
        lngCol = .Col
        .Cell(flexcpBackColor, lngRow, lngCol) = HighlightBackColor
        .Cell(flexcpForeColor, lngRow, lngCol) = HighlightForeColor
        If frmRunLimitTimeEdit.ShowMe(0, Mid(tvwPlanTree.SelectedItem.Key, 3), lngRow, strTimeStart, strTimeStop) Then
            Call tvwPlanTree_NodeClick(tvwPlanTree.SelectedItem)
            If lngRow > .Rows - 1 Then lngRow = .Rows - 1
            If lngCol > .Cols - 1 Then lngCol = .Cols - 1
            .Row = lngRow
            .Col = lngCol
        End If
        .Cell(flexcpBackColor, lngRow, lngCol) = &HFFF6D2
        .Cell(flexcpForeColor, lngRow, lngCol) = &H80000008
    End With
End Sub

Private Sub DeleteTime()
'ɾ��ʱ���
    Dim strTimeStart As String, strTimeStop As String
    Dim lngRow As Long, lngCol As Long
    
    On Error GoTo errH
    If tvwPlanTree.SelectedItem Is Nothing Then Exit Sub
    If tvwPlanTree.SelectedItem.Text = "Ԥ�跽��" Then Exit Sub
    With vsfPlanDetail
        If .Row = 0 Or .Row = 8 Or .Col = 0 Or .TextMatrix(.Row, .Col) = "" Then Exit Sub
        strTimeStart = Mid(Split(.TextMatrix(.Row, .Col), vbNewLine)(0), 3)
        strTimeStop = Mid(Split(.TextMatrix(.Row, .Col), vbNewLine)(2), 3)
        lngRow = .Row
        lngCol = .Col
        If MsgBox("ȷ��Ҫ��ʱ��Ρ�" & strTimeStart & "-" & strTimeStop & "��ɾ����", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
            Call ExecuteProcedure("Zl_ZlRunLimitTime_Update(2," & .Cell(flexcpData, .Row, .Col) & ")", "ɾ��ʱ���")
            Call tvwPlanTree_NodeClick(tvwPlanTree.SelectedItem)
            If lngRow > .Rows - 1 Then lngRow = .Rows - 1
            If lngCol > .Cols - 1 Then lngCol = .Cols - 1
            .Row = lngRow
            .Col = lngCol
        End If
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub EditTime()
'�޸�ʱ���
    Dim strTimeStart As String, strTimeStop As String
    Dim lngRow As Long, lngCol As Long
    
    If tvwPlanTree.SelectedItem Is Nothing Then Exit Sub
    If tvwPlanTree.SelectedItem.Text = "Ԥ�跽��" Then Exit Sub
    With vsfPlanDetail
        If .Row = 0 Or .Col = 0 Or .TextMatrix(.Row, .Col) = "" Then Exit Sub
        strTimeStart = Mid(Split(.TextMatrix(.Row, .Col), vbNewLine)(0), 3)
        strTimeStop = Mid(Split(.TextMatrix(.Row, .Col), vbNewLine)(2), 3)
        lngRow = .Row
        lngCol = .Col
        .Cell(flexcpBackColor, lngRow, lngCol) = HighlightBackColor
        .Cell(flexcpForeColor, lngRow, lngCol) = HighlightForeColor
        If frmRunLimitTimeEdit.ShowMe(.Cell(flexcpData, .Row, .Col), Mid(tvwPlanTree.SelectedItem.Key, 3), .Row, strTimeStart, strTimeStop) Then
            Call tvwPlanTree_NodeClick(tvwPlanTree.SelectedItem)
            If lngRow > .Rows - 1 Then lngRow = .Rows - 1
            If lngCol > .Cols - 1 Then lngCol = .Cols - 1
            .Row = lngRow
            .Col = lngCol
        End If
        .Cell(flexcpBackColor, lngRow, lngCol) = &HFFF6D2
        .Cell(flexcpForeColor, lngRow, lngCol) = &H80000008
    End With
End Sub

Private Sub NewPlan()
'��������
    Dim lngPlanNo As Long
    Dim objNode As Node
    Dim strPlanName As String, strDescription As String
    
    If frmRunLimitPlanEdit.ShowMe(Me, lngPlanNo, strPlanName, strDescription) Then
        Set objNode = tvwPlanTree.Nodes.Add(, , "K_" & lngPlanNo, strPlanName, "enabled")
        objNode.Tag = strDescription
        tvwPlanTree.Nodes("K_" & lngPlanNo).Selected = True
        Call tvwPlanTree_NodeClick(tvwPlanTree.SelectedItem)
    End If
End Sub

Private Sub RemovePlan()
'ɾ������
    On Error GoTo errH
    If tvwPlanTree.SelectedItem Is Nothing Then Exit Sub
    '�ж��Ƿ���ģ������ʹ�ø÷�������������ʾ
    If CheckPlanStatus("ɾ��") = False Then Exit Sub
    
    If MsgBox("ȷ��Ҫɾ���÷�����", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) Then
        Call ExecuteProcedure("Zl_Zlrunlimit_Update(2," & mlngPlanNo & ")", "ɾ������")
        tvwPlanTree.Nodes.Remove (tvwPlanTree.SelectedItem.Key)
        If tvwPlanTree.Nodes.Count > 0 Then
            tvwPlanTree.Tag = ""
            Call tvwPlanTree_NodeClick(tvwPlanTree.SelectedItem)
        Else
            Call SetEnabled(False)
            Call ClearPlanDetail
        End If
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub StartPlan()
'���÷���
    On Error GoTo errH
    If tvwPlanTree.SelectedItem Is Nothing Then Exit Sub
    If tvwPlanTree.SelectedItem.Image = "enabled" Or tvwPlanTree.SelectedItem.Image = "enabledLock" Then Exit Sub
    Call ExecuteProcedure("Zl_Zlrunlimit_Update(1," & mlngPlanNo & ",Null,1)", "���÷���")
    If tvwPlanTree.SelectedItem.Image = "disabled" Then
        tvwPlanTree.SelectedItem.Image = "enabled"
    Else
        tvwPlanTree.SelectedItem.Image = "enabledLock"
    End If
    Call tvwPlanTree_NodeClick(tvwPlanTree.SelectedItem)
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub StopPlan()
'ͣ�÷���
    On Error GoTo errH
    If tvwPlanTree.SelectedItem Is Nothing Then Exit Sub
    If tvwPlanTree.SelectedItem.Image = "disabled" Or tvwPlanTree.SelectedItem.Image = "disabledLock" Then Exit Sub
    '�ж��Ƿ���ģ������ʹ�ø÷�������������ʾ
    If CheckPlanStatus("ͣ��") = False Then Exit Sub
    Call ExecuteProcedure("Zl_Zlrunlimit_Update(1," & mlngPlanNo & ",Null,0)", "ͣ�÷���")
    If tvwPlanTree.SelectedItem.Image = "enabled" Then
        tvwPlanTree.SelectedItem.Image = "disabled"
    Else
        tvwPlanTree.SelectedItem.Image = "disabledLock"
    End If
    If mnuViewShow.Checked = False Then
        tvwPlanTree.Nodes.Remove (tvwPlanTree.SelectedItem.Key)
        If tvwPlanTree.Nodes.Count > 0 Then
            tvwPlanTree.Tag = ""
            Call tvwPlanTree_NodeClick(tvwPlanTree.SelectedItem)
        Else
            Call SetEnabled(False)
            Call ClearPlanDetail
        End If
    Else
        Call tvwPlanTree_NodeClick(tvwPlanTree.SelectedItem)
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'��鷽���Ƿ����ڱ�ʹ��
Private Function CheckPlanStatus(ByVal strTag As String) As Boolean
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strFuncList As String
    Dim i As Long
    
    On Error GoTo errH
    strSql = "Select a.���� From Zlrunlimitset A, Zlrunlimit B Where a.������� = b.��� And b.��� =[1]"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, Me.Caption, mlngPlanNo)
    With rsTemp
        If .RecordCount > 0 Then
            For i = 1 To .RecordCount
                If i > 3 Then Exit For
                strFuncList = strFuncList & "��" & !���� & "��" & vbNewLine
                .MoveNext
            Next
            If .RecordCount > 3 Then
                MsgBox "����ʱ�������ڱ�" & vbNewLine & strFuncList & "��" & .RecordCount & _
                "������ʹ�ã�Ҫ" & strTag & "�÷��������޸����Ϲ��ܵķ�����", vbInformation, gstrSysName
            Else
                MsgBox "����ʱ�������ڱ�����" & vbNewLine & strFuncList & _
                "ʹ�ã�Ҫ" & strTag & "�÷��������޸����Ϲ��ܵķ�����", vbInformation, gstrSysName
            End If
            Exit Function
        End If
    End With
    CheckPlanStatus = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub UpdatePlan()
'�޸ķ���
    Dim strPlanName As String, strDescription As String
    
    If tvwPlanTree.SelectedItem Is Nothing Then Exit Sub
    If tvwPlanTree.SelectedItem.Text = "Ԥ�跽��" Then Exit Sub
    strPlanName = tvwPlanTree.SelectedItem.Text
    strDescription = tvwPlanTree.SelectedItem.Tag
    If frmRunLimitPlanEdit.ShowMe(Me, mlngPlanNo, strPlanName, strDescription) Then
        tvwPlanTree.Nodes("K_" & mlngPlanNo).Text = strPlanName
        tvwPlanTree.Nodes("K_" & mlngPlanNo).Tag = strDescription
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuPlanDetailAdd_Click()
    Call AddTime
End Sub

Private Sub mnuPlanDetailDel_Click()
    Call DeleteTime
End Sub

Private Sub mnuPlanDetailModify_Click()
    Call EditTime
End Sub

Private Sub mnuPlanNameNew_Click()
    Call NewPlan
End Sub

Private Sub mnuPlanNameRemove_Click()
    Call RemovePlan
End Sub

Private Sub mnuPlanNameStart_Click()
    Call StartPlan
End Sub

Private Sub mnuPlanNameStop_Click()
    Call StopPlan
End Sub

Private Sub mnuPlanNameUpdate_Click()
    Call UpdatePlan
End Sub

Private Sub ShowStopPlan()
    mnuViewShow.Checked = Not mnuViewShow.Checked
    If tvwPlanTree.Nodes.Count > 0 Then
        mlngPlanNo = Split(tvwPlanTree.Nodes(1).Key, "_")(1)
    End If
    Call FillPlanList
End Sub

Private Sub tvwPlanTree_DblClick()
    Call UpdatePlan
End Sub

Private Sub tvwPlanTree_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPlanName
    End If
End Sub

Private Sub tvwPlanTree_NodeClick(ByVal Node As MSComctlLib.Node)
    If tvwPlanTree.Tag <> "" Then
        tvwPlanTree.Nodes(tvwPlanTree.Tag).BackColor = &H80000005
        tvwPlanTree.Nodes(tvwPlanTree.Tag).ForeColor = &H80000012
    End If
    Node.BackColor = HighlightBackColor
    Node.ForeColor = HighlightForeColor
    tvwPlanTree.Tag = tvwPlanTree.SelectedItem.Key
    mlngPlanNo = Split(Node.Key, "_")(1)
    If Node.Text = "Ԥ�跽��" Then
        Call SetEnabled(False)
    Else
        Call SetEnabled(True)
    End If
    Call FillPlanDetail
End Sub

Private Sub SetEnabled(ByVal blnEnabled As Boolean)
    mnuPlanNameUpdate.Enabled = blnEnabled
    mnuPlanNameRemove.Enabled = blnEnabled
    mnuPlanDetailAdd.Enabled = blnEnabled
    mnuPlanDetailModify.Enabled = blnEnabled
    mnuPlanDetailDel.Enabled = blnEnabled
    If tvwPlanTree.Nodes.Count = 0 Then
        mnuPlanNameStart.Enabled = False
        mnuPlanNameStop.Enabled = False
    Else
        If tvwPlanTree.SelectedItem.Image = "enabledLock" Or tvwPlanTree.SelectedItem.Image = "enabled" Then
            mnuPlanNameStart.Enabled = False
            mnuPlanNameStop.Enabled = True
        Else
            mnuPlanNameStart.Enabled = True
            mnuPlanNameStop.Enabled = False
        End If
    End If
End Sub

Private Sub FillPlanDetail()
'�����ϸ������Ϣ
    Dim j As Long  '��ʾʱ���
    Dim lngLastWeekNo As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
        '���ϵķ�����Ϣ���
        Call ClearPlanDetail
        
        '����·���
        strSql = "Select Id, ����, To_Char(��ʼʱ��, 'HH24:MI:SS') ��ʼʱ��, To_Char(����ʱ��, 'HH24:MI:SS') ����ʱ��" & vbNewLine & _
                "From ZlRunLimitTime" & vbNewLine & _
                "Where ���� = [1]" & vbNewLine & _
                "Order By ����, ��ʼʱ��, ����ʱ��"
        Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, Me.Caption, mlngPlanNo)
        With rsTemp
            Do While Not .EOF
                If !���� = lngLastWeekNo Then
                    j = j + 1
                    If j + 2 > vsfPlanDetail.Cols Then
                        vsfPlanDetail.Cols = j + 2
                        vsfPlanDetail.ColWidth(j) = vsfPlanDetail.ColWidth(PDT_ʱ���1)
                        vsfPlanDetail.TextMatrix(0, j) = "ʱ���" & j
                        vsfPlanDetail.ColAlignment(j) = flexAlignCenterCenter
                        Call AdjustFormDisplay
                    End If
                Else
                    j = 1
                End If
                vsfPlanDetail.TextMatrix(!���� + 1, j) = "�� " & !��ʼʱ�� & vbNewLine & vbNewLine & "ֹ " & !����ʱ��
                vsfPlanDetail.Cell(flexcpData, !���� + 1, j) = Val(!id & "")
                lngLastWeekNo = !����
                .MoveNext
            Loop
        End With
        vsfPlanDetail.ToolTipText = tvwPlanTree.SelectedItem.Tag
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'���ϵķ�����Ϣ���
Private Sub ClearPlanDetail()
    Dim i As Long
    
    With vsfPlanDetail
        .Cols = 3
        .TextMatrix(0, PDT_ʱ�����չ) = ""
        For i = 1 To 7
            .TextMatrix(i, PDT_ʱ���1) = ""
            .TextMatrix(i, PDT_ʱ�����չ) = ""
        Next
        Call AdjustFormDisplay
    End With
End Sub

Private Sub vsfPlanDetail_DblClick()
    With vsfPlanDetail
        If .MouseRow <> .Row Then Exit Sub
        If .TextMatrix(.Row, .Col) = "" Then
            Call AddTime
        Else
            Call EditTime
        End If
    End With
End Sub

Private Sub vsfPlanDetail_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDelete
        Call DeleteTime
    End Select
End Sub

Private Sub vsfPlanDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPlanDetail
    End If
End Sub
