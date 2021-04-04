VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmFeedBackList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "疾病报告和阳性反馈单对应"
   ClientHeight    =   9570
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10530
   ControlBox      =   0   'False
   Icon            =   "frmFeedBackList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ListBox lbList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   840
      Width           =   10245
   End
   Begin VB.PictureBox picReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4065
      ScaleWidth      =   10245
      TabIndex        =   2
      Top             =   1710
      Width           =   10275
      Begin XtremeSuiteControls.TabControl tbcReport 
         Height          =   3255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   9735
         _Version        =   589884
         _ExtentX        =   17171
         _ExtentY        =   5741
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picWarn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   720
      ScaleHeight     =   5895
      ScaleWidth      =   8895
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   8895
      Begin VB.Image Image1 
         Height          =   1245
         Left            =   960
         Picture         =   "frmFeedBackList.frx":6852
         Top             =   1800
         Width           =   1500
      End
      Begin VB.Label lblWarn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "该病人没有填写反馈单！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   2760
         TabIndex        =   1
         Top             =   2160
         Width           =   5535
      End
   End
   Begin VB.Label lblFeedBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "请选择该报告对应的阳性结果反馈单："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   555
      Width           =   3735
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmFeedBackList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsDisease As ADODB.Recordset     '该报告可能对应的阳性结果数据集
Private mfrmReport As frmDiseaseRegist
Private mblnResult As Boolean
Private mstrIDs As String '反馈单ID
Private mintType As Integer        '1-和报告卡关联；2-选择一张反馈单

Public Function ShowMe(ByVal frmParent As Object, ByVal rsDis As ADODB.Recordset, ByRef strIDs As String, Optional ByVal intType As Integer = 1) As Boolean
On Error GoTo errHand
    Set mrsDisease = rsDis
    mintType = intType
    Me.Show 1, frmParent
    strIDs = mstrIDs
    ShowMe = mblnResult
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetSelectedIDs()
    Dim lngCount As Long
    mstrIDs = ""
    For lngCount = 0 To lbList.ListCount - 1
        If lbList.Selected(lngCount) Then
            mstrIDs = mstrIDs & "," & lbList.ItemData(lngCount)
        End If
    Next
    If mstrIDs <> "" Then mstrIDs = Mid(mstrIDs, 2)
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Tool_OK
            Call GetSelectedIDs
            If mintType = 1 Then
                If mstrIDs = "" Then
                    If MsgBox("您确定一张反馈单都不关联吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
            ElseIf mintType = 2 Then
                If mstrIDs = "" Then
                    MsgBox "必须选择一张反馈单！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            mblnResult = True
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    If mintType = 1 Then
        Me.Caption = "疾病报告和阳性反馈单对应"
        lblFeedBack.Caption = "请选择该报告对应的阳性结果反馈单："
    ElseIf mintType = 2 Then
        Me.Caption = "选择阳性反馈单"
        lblFeedBack.Caption = "请选择一张反馈单："
    End If
    Call InitCommandBar
    Set mfrmReport = New frmDiseaseRegist
    Call mfrmReport.SetFrmInset(False)

    Call InitTabContol
    Call InitDiseaseList
    Call picReport_Resize
    Call lbList_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Me.Height = 11500
    Me.Top = (VB.Screen.Height - Me.Height) / 2
    Call picReport_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmReport Is Nothing Then
        Unload mfrmReport
        Set mfrmReport = Nothing
    End If
    Set mrsDisease = Nothing
End Sub

Private Sub lbList_Click()
    Dim lngID As Long
    
    lngID = Val(lbList.ItemData(lbList.ListIndex))
    If lngID <> 0 Then
        Call mfrmReport.zlRefresh(lngID)
    End If
End Sub

Private Sub lbList_ItemCheck(Item As Integer)
'功能：设置只能够选择一张反馈单
    Dim lngCount As Long
    If mintType = 2 Then
        For lngCount = 0 To lbList.ListCount - 1
            If lbList.Selected(lngCount) And lngCount <> Item Then
                lbList.Selected(lngCount) = False
            End If
        Next
    End If
End Sub

Private Sub picReport_Resize()
On Error Resume Next
    picReport.Height = 9300
    tbcReport.Move picReport.ScaleLeft, picReport.ScaleTop, picReport.ScaleWidth, picReport.ScaleHeight
End Sub

Private Sub InitTabContol()
    With Me.tbcReport
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
            .ButtonMargin.SetRect -20, -20, 0, 0
            .HeaderMargin.SetRect 0, 0, 0, 0
        End With

        Call .InsertItem(0, "反馈单", mfrmReport.hwnd, 0)
        Call .InsertItem(1, "无反馈单", picWarn.hwnd, 0)
        .Item(1).Selected = True
        .Item(0).Selected = True
        .Item(1).Visible = False
    End With
End Sub
    
Private Sub InitDiseaseList()
On Error GoTo errH:
    With mrsDisease
        lbList.Clear
        Do While Not .EOF
            lbList.AddItem !NO & "." & !科室 & "(登记时间:" & !登记时间 & ")"
            lbList.ItemData(Me.lbList.NewIndex) = !ID
            If glngOpenedID <> 0 Then
                If glngOpenedID = !ID Then
                    lbList.Selected(lbList.NewIndex) = True
                    lbList.ListIndex = lbList.NewIndex
                End If
            End If
            .MoveNext
        Loop
        If lbList.ListIndex = -1 Then lbList.ListIndex = 0
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lbList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call gobjComlib.ZLCommFun.PressKey(vbKeyTab)
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox

    '工具栏----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = gobjComlib.ZLCommFun.GetPubIcons
    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Tool_OK, "确定"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
End Sub

