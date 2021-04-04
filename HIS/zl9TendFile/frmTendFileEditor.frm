VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendFileEditor 
   BackColor       =   &H00FFFFFF&
   Caption         =   "护理记录查阅"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11415
   Icon            =   "frmTendFileEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picPrompt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1410
      ScaleHeight     =   285
      ScaleWidth      =   9645
      TabIndex        =   2
      Top             =   7710
      Width           =   9645
      Begin VB.Label lblPrompt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   30
         TabIndex        =   3
         Top             =   60
         Width           =   10500
      End
   End
   Begin zl9TendFile.usrTendFileEditor usrTendFileEditor 
      Height          =   6045
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10663
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7560
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTendFileEditor.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17224
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmTendFileEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
'窗体变量
'######################################################################################################################
Public mblnDoctorStation As Boolean
Public mstrPrivs As String
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mint婴儿 As Integer
Private mlngFileID As Long
Private mblnChildForm As Boolean
Private mblnStartUp As Boolean
Private mblnEdit As Boolean
Private mblnChange As Boolean                           '修改标志
Private mblnSign As Boolean                             '签名标志
Private mblnArchive As Boolean                          '归档标志
Private mblnOK As Boolean                               '数据是否已经发生变化
Private mfrmTipInfo As Object

Public WithEvents zlEvent_Print As zlTFPrintMethod
Attribute zlEvent_Print.VB_VarHelpID = -1
Public Event zlAfterPrint(ByVal lngFileID As Long)
Public Event AfterDataChanged(ByVal blnChange As Boolean)
Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)

Private mbytFontSize As Byte

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-19 15:16
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
    Call usrTendFileEditor.SetFontSize(bytSize)
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-20 15:15:00
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytFontSize As Byte
    bytFontSize = mbytFontSize
    
    Me.FontSize = bytFontSize
    Me.FontName = "宋体"
    
    lblPrompt.FontSize = bytFontSize
    
    Set CtlFont = cbsMain.Options.Font
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = bytFontSize
    Set cbsMain.Options.Font = CtlFont
    cbsMain.RecalcLayout
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Save  '保存
        Call SaveData
    Case conMenu_Edit_Transf_Cancle  '取消
        Call CancelData
    Case conMenu_Tool_Sign  '签名
        Call SignData(False, False)
    Case conMenu_Tool_SignShiftExchange '交班签名
        Call SignData(False, True)
    Case conMenu_Tool_SignEarse  '取消签名
        Call UnSignData(False)
    Case conMenu_Tool_SignAuditAffirm '审签
        Call SignData(True, False)
    Case conMenu_Tool_SignAuditCancel '取消审签
        Call UnSignData(True)
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Call Form_Resize
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Save  '保存
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "护理记录登记") > 0)
        Control.Enabled = mblnChange And Not mblnArchive And Control.Visible
    Case conMenu_Edit_Transf_Cancle  '取消
        Control.Visible = Not mblnDoctorStation And Not gblnMoved
        Control.Enabled = mblnChange And Not mblnArchive And Control.Visible
    Case conMenu_Tool_Sign, conMenu_Tool_SignShiftExchange '签名
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "护理记录签名") > 0)
        Control.Enabled = Not mblnArchive And Not mblnChange And Control.Visible And mblnEdit
    Case conMenu_Tool_SignEarse '取消签名
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "取消记录签名") > 0)
        Control.Enabled = Not mblnArchive And Not mblnChange And Control.Visible And mblnEdit
    Case conMenu_Tool_SignAuditAffirm, conMenu_Tool_SignAuditCancel '审签,取消审签
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "护理记录审签") > 0)
        Control.Enabled = Not mblnArchive And Not mblnChange And Control.Visible And mblnEdit
        If Control.ID = conMenu_Tool_SignAuditCancel And Control.Enabled Then
            Control.Enabled = (InStr(1, mstrPrivs, "取消记录签名") > 0)
        End If
    End Select
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    If mfrmTipInfo Is Nothing Then
        Set mfrmTipInfo = New frmTipInfo
    End If
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call Me.cbsMain.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    If Me.WindowState = vbMinimized Then Exit Sub
    Err = 0: On Error Resume Next
    With usrTendFileEditor
        .Left = 0
        .Top = lngScaleTop
        .Width = lngScaleRight
        .Height = lngScaleBottom - lngScaleTop - IIf(stbThis.Visible, stbThis.Height, 0)
    End With

    With picPrompt
        .Visible = stbThis.Visible
        .Top = stbThis.Top + 50
        .Height = stbThis.Height - 100
        .Left = stbThis.Panels(2).Left + 50
        .Width = stbThis.Panels(2).Width - 100
    End With
    With lblPrompt
        .Visible = stbThis.Visible
        .Width = picPrompt.Width
        .Height = TextHeight("刘")
        .Top = (picPrompt.Height - .Height) \ 2
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmTipInfo Is Nothing Then
        Unload mfrmTipInfo
        Set mfrmTipInfo = Nothing
    End If
    Set zlEvent_Print = Nothing
    If mblnChildForm = False Then Call SaveWinState(Me, App.ProductName)
End Sub

Public Function ShowMe(ByVal frmParent As Form, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, _
    ByVal lngDeptID As Long, ByVal intBaby As Integer, ByVal blnChildForm As Boolean, ByVal strPrivs As String, _
    Optional ByVal blnEdit As Boolean, Optional ByVal bytSize As Byte = 0) As Boolean
    '******************************************************************************************************************
    '功能： 显示护理记录文件内容
    '参数： frmParent           上级窗体对象
    '       lngFileID           护理文件格式句柄
    '       lngPatiID           病人id
    '       lngPageID           主页id
    '       intBaby             婴儿标志
    '返回： 无
    '******************************************************************************************************************
'    Dim bln护理级别 As Boolean
    
    Err = 0
    On Error GoTo errHand
    
    mlngFileID = lngFileID
    mblnChildForm = blnChildForm
    mlng病人ID = lngPatiID
    mlng主页ID = lngPageId
    mint婴儿 = intBaby
    mblnEdit = blnEdit
    mblnOK = False
    
    If mblnChildForm Then
        If mblnStartUp Then
            Call FormSetCaption(Me, False, False)
            mblnStartUp = False
        End If
    Else
        Me.WindowState = 2
'        blnEdit = False
        
        Call MainDefMenus
    End If
    stbThis.Visible = Not mblnChildForm
    cbsMain.ActiveMenuBar.Visible = False
    cbsMain.RecalcLayout
    
    Call usrTendFileEditor.ShowMe(Me, lngFileID, lngPatiID, lngPageId, lngDeptID, intBaby, strPrivs, blnEdit)
    
    '窗体显示
    If blnChildForm = False Then
        Call SetFontSize(bytSize)
        If frmParent Is Nothing Then
            Me.Show vbModal
        Else
            Me.Show vbModal, frmParent
        End If
        ShowMe = mblnOK
        Unload Me
        Exit Function
    End If
    
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub MainDefMenus()
    Dim cbrControl As Object
    Dim cbrToolBar As Object
    
    CommandBarsGlobalSettings.App = App
    Set Me.cbsMain.Icons = zlCommFun.GetPubIcons
    With Me.cbsMain.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With
    
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "取消")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "签名"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignShiftExchange, "交班签名")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "回退"): cbrControl.IconId = 229
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignAuditAffirm, "审签"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignAuditCancel, "回退")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): cbrControl.BeginGroup = True
    End With
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
End Sub

Private Sub usrTendFileEditor_AfterDataChanged(ByVal blnChange As Boolean)
    mblnChange = blnChange
    RaiseEvent AfterDataChanged(blnChange)
End Sub

Private Sub usrTendFileEditor_AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)
    mblnArchive = blnArchive
    mblnSign = blnSign
    lblPrompt.Caption = strInfo
    lblPrompt.ForeColor = IIf(blnImportant, &HFF&, &H80000008)
    
    RaiseEvent AfterRowColChange(strInfo, blnImportant, blnSign, blnArchive)
End Sub

Private Sub usrTendFileEditor_ShowTipInfo(ByVal vsfObj As Object, ByVal strInfo As String, ByVal blnMultiRow As Boolean)
    Call mfrmTipInfo.ShowTipInfo(vsfObj, strInfo, blnMultiRow)
End Sub

Private Sub zlEvent_Print_zlAfterPrint()
    RaiseEvent zlAfterPrint(mlngFileID)
End Sub

Public Function SaveData() As Boolean
    SaveData = usrTendFileEditor.SaveME
    If mblnOK = False Then mblnOK = SaveData
End Function

Public Function CancelData() As Boolean
    CancelData = usrTendFileEditor.CancelMe
End Function

Public Sub SignData(blnVerify As Boolean, blnExchange As Boolean)
    Call usrTendFileEditor.SignMe(blnVerify, blnExchange)
    mblnOK = True
End Sub

Public Sub UnSignData(blnVerify As Boolean)
    Call usrTendFileEditor.UnSignMe(blnVerify)
    mblnOK = True
End Sub

Public Sub ArchiveData()
    Call usrTendFileEditor.ArchiveMe
    mblnOK = True
End Sub

Public Sub UnArchiveData()
    Call usrTendFileEditor.UnArchiveMe
    mblnOK = True
End Sub

Public Sub SignMarker()
    Call usrTendFileEditor.SignMarker
End Sub

Public Sub SetArchiveData(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intBaby As Integer)
    Call usrTendFileEditor.SetArchiveValue(lngPatiID, lngPageId, intBaby)
End Sub

Public Function zlPrintTend(Optional ByVal bytMode As Byte = 2, Optional ByVal strPrintDeviceName As String) As Boolean
    '1-预览,2-打印
    
    Select Case bytMode
    Case 1
        Call zlRptPrint(2, strPrintDeviceName)
    Case 2
        Call zlRptPrint(1, strPrintDeviceName)
    Case 3
        Call zlRptPrint(3, strPrintDeviceName)
    End Select
End Function

Private Sub zlRptPrint(ByVal bytMode As Byte, Optional ByVal strPrintDeviceName As String)
    Dim objPrint As New zlTFPrintTends, objAppRow As zlTFTabAppRow
    Dim lngWidth As Long
    
    If zlEvent_Print Is Nothing Then
        Set zlEvent_Print = New zlTFPrintMethod
    End If
    
    Call zlEvent_Print.InitPrint(gcnOracle, gstrDBUser)
    If strPrintDeviceName = "" Then
        bytMode = zlEvent_Print.zlPrintAsk(mlng病人ID, mlng主页ID, mint婴儿, mlngFileID)
    Else
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "DeviceName", strPrintDeviceName
        Call zlEvent_Print.zlPrintAsk(mlng病人ID, mlng主页ID, mint婴儿, mlngFileID, True)
    End If
    
    If bytMode <> 0 Then zlEvent_Print.zlPrintOrViewTends (strPrintDeviceName <> ""), bytMode
End Sub
