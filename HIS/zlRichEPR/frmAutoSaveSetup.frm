VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAutoSaveSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "自动化设置"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   5400
   Icon            =   "frmAutoSaveSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkPrintNoAsk 
      Caption         =   "静默打印(&S)"
      Height          =   285
      Left            =   3885
      TabIndex        =   21
      Top             =   2220
      Width           =   1305
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   855
      TabIndex        =   9
      ToolTipText     =   "范围：1～100"
      Top             =   2595
      Width           =   615
   End
   Begin VB.CheckBox chkAutoPageNote 
      Caption         =   "内容增页提醒(&N)"
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   2220
      Width           =   1695
   End
   Begin VB.CheckBox chkAutoPageCount 
      Caption         =   "自动分页计数(&P)"
      Height          =   285
      Left            =   270
      TabIndex        =   6
      Top             =   2220
      Width           =   1860
   End
   Begin MSComCtl2.UpDown udIntervalEPR 
      Height          =   285
      Left            =   2820
      TabIndex        =   5
      Top             =   1815
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "txtAutoSaveEPR"
      BuddyDispid     =   196614
      OrigLeft        =   3330
      OrigTop         =   1845
      OrigRight       =   3585
      OrigBottom      =   2130
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtAutoSaveEPR 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      ToolTipText     =   "范围：1～100"
      Top             =   1815
      Width           =   780
   End
   Begin VB.CheckBox chkAutoSaveEPR 
      Caption         =   "定时自动保存(&A)"
      Height          =   285
      Left            =   270
      TabIndex        =   3
      Top             =   1815
      Width           =   1740
   End
   Begin VB.CheckBox chkAutoSave 
      Caption         =   "开启“自动缓存”(&T)"
      Height          =   285
      Left            =   270
      TabIndex        =   0
      Top             =   105
      Width           =   2400
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2820
      TabIndex        =   11
      Top             =   3000
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3915
      TabIndex        =   12
      Top             =   3000
      Width           =   1100
   End
   Begin VB.Frame fraPati 
      Caption         =   " 自动缓存设置 "
      Height          =   1215
      Left            =   555
      TabIndex        =   15
      Top             =   435
      Width           =   4635
      Begin MSComCtl2.UpDown udMax 
         Height          =   285
         Left            =   3226
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   675
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMax"
         BuddyDispid     =   196620
         OrigLeft        =   3555
         OrigTop         =   720
         OrigRight       =   3810
         OrigBottom      =   960
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtMax 
         Height          =   285
         Left            =   2430
         TabIndex        =   2
         ToolTipText     =   "范围：1～100"
         Top             =   675
         Width           =   1050
      End
      Begin MSComCtl2.UpDown udInterval 
         Height          =   285
         Left            =   3226
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   293
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtInterval"
         BuddyDispid     =   196621
         OrigLeft        =   3465
         OrigTop         =   315
         OrigRight       =   3720
         OrigBottom      =   600
         Max             =   9999
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtInterval 
         Height          =   285
         Left            =   2430
         TabIndex        =   1
         ToolTipText     =   "范围：1～9999"
         Top             =   293
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "步"
         Height          =   195
         Left            =   3555
         TabIndex        =   19
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "撤销/重做的最大步数(&R)"
         Height          =   195
         Left            =   225
         TabIndex        =   18
         Top             =   720
         Width           =   2040
      End
      Begin VB.Label Label2 
         Caption         =   "秒"
         Height          =   195
         Left            =   3555
         TabIndex        =   17
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "自动缓存时间间隔(&P)"
         Height          =   195
         Left            =   225
         TabIndex        =   16
         Top             =   360
         Width           =   1860
      End
   End
   Begin MSComCtl2.UpDown udn 
      Height          =   285
      Left            =   1485
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2595
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txt"
      BuddyDispid     =   196610
      OrigLeft        =   3555
      OrigTop         =   720
      OrigRight       =   3810
      OrigBottom      =   960
      Max             =   100
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "显示前           次的共享页面文件历史内容(&B)"
      Height          =   180
      Left            =   255
      TabIndex        =   8
      Top             =   2640
      Width           =   3960
   End
   Begin VB.Label Label5 
      Caption         =   "分钟"
      Height          =   195
      Left            =   3165
      TabIndex        =   20
      Top             =   1860
      Width           =   600
   End
End
Attribute VB_Name = "frmAutoSaveSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnAutosave As Boolean         '是否开启自动缓存功能
Private mlngUndoLimit As Long           '自动缓存步数
Private mlngSaveInterval As Long        '时间间隔
Private mblnAutoSaveEPR As Boolean      '是否开启自动保存
Private mlngSaveIntervalEPR As Long     '自动保存的时间间隔
Private mblnAutoPageCount As Boolean                    '自动分页计数
Private mblnAutoPageNote As Boolean                     '自动增页提醒
Private mintSharePages As Integer       '显示历史内容的次数
Private mblnNoAsk As Boolean            '静默打印

Private mblnOk As Boolean, mstrPrivs As String

Public Function ShowMe(ByRef frmParent As Object, ByVal strPrivs As String) As Boolean
    mstrPrivs = strPrivs
    cmdOK.Enabled = InStr(mstrPrivs, "参数设置") > 0
    mblnAutosave = zlDatabase.GetPara("AutoSave", glngSys, 1070, 1, Array(chkAutoSave), InStr(mstrPrivs, "参数设置") > 0) = 1
    mlngUndoLimit = zlDatabase.GetPara("UndoLimit", glngSys, 1070, 20, Array(txtMax, udMax), InStr(mstrPrivs, "参数设置") > 0)
    mlngSaveInterval = zlDatabase.GetPara("SaveInterval", glngSys, 1070, 60, Array(txtInterval, udInterval), InStr(mstrPrivs, "参数设置") > 0)
    mblnAutoSaveEPR = zlDatabase.GetPara("AutoSaveEPR", glngSys, 1070, 0, Array(chkAutoSaveEPR), InStr(mstrPrivs, "参数设置") > 0) = 1
    mlngSaveIntervalEPR = zlDatabase.GetPara("SaveIntervalEPR", glngSys, 1070, 5, Array(txtAutoSaveEPR, udIntervalEPR), InStr(mstrPrivs, "参数设置") > 0)
    mblnAutoPageCount = zlDatabase.GetPara("AutoPageCount", glngSys, 1070, 0, Array(chkAutoPageCount), InStr(mstrPrivs, "参数设置") > 0) = 1
    mblnAutoPageNote = zlDatabase.GetPara("AutoPageNote", glngSys, 1070, 0, Array(chkAutoPageNote), InStr(mstrPrivs, "参数设置") > 0) = 1
    mintSharePages = zlDatabase.GetPara("SharePageCount", glngSys, 1070, 5, Array(txt, udn), InStr(mstrPrivs, "参数设置") > 0)
    mblnNoAsk = zlDatabase.GetPara("NoAsk", glngSys, 1070, 0, Array(chkPrintNoAsk), InStr(mstrPrivs, "参数设置") > 0) = 1
    
    '设置显示状态
    chkAutoSave.Value = IIf(mblnAutosave, vbChecked, vbUnchecked)
    txtMax = mlngUndoLimit
    txtInterval = mlngSaveInterval
    chkAutoSaveEPR.Value = IIf(mblnAutoSaveEPR, vbChecked, vbUnchecked)
    txtAutoSaveEPR = mlngSaveIntervalEPR
    chkAutoPageCount.Value = IIf(mblnAutoPageCount, vbChecked, vbUnchecked)
    chkAutoPageNote.Value = IIf(mblnAutoPageNote, vbChecked, vbUnchecked)
    txt.Text = mintSharePages
    chkAutoSave.Value = IIf(mblnAutosave, vbChecked, vbUnchecked)
    chkPrintNoAsk.Value = IIf(mblnNoAsk, vbChecked, vbUnchecked)
    
    Call chkAutoSave_Click
    Call chkAutoSaveEPR_Click
    Me.Show vbModal, frmParent
    If mblnOk Then
        zlDatabase.SetPara "AutoSave", IIf(mblnAutosave, 1, 0), glngSys, 1070
        zlDatabase.SetPara "UndoLimit", mlngUndoLimit, glngSys, 1070
        zlDatabase.SetPara "SaveInterval", mlngSaveInterval, glngSys, 1070
        zlDatabase.SetPara "AutoSaveEPR", IIf(mblnAutoSaveEPR, 1, 0), glngSys, 1070
        zlDatabase.SetPara "SaveIntervalEPR", mlngSaveIntervalEPR, glngSys, 1070
        zlDatabase.SetPara "AutoPageCount", IIf(mblnAutoPageCount, 1, 0), glngSys, 1070
        zlDatabase.SetPara "AutoPageNote", IIf(mblnAutoPageNote, 1, 0), glngSys, 1070
        zlDatabase.SetPara "SharePageCount", mintSharePages, glngSys, 1070
        zlDatabase.SetPara "NoAsk", IIf(mblnNoAsk, 1, 0), glngSys, 1070
    End If
    ShowMe = mblnOk
End Function

Private Sub chkAutoPageCount_Click()
    If Me.chkAutoPageCount.Value = vbChecked Then
        Me.chkAutoPageNote.Enabled = True
    Else
        Me.chkAutoPageNote.Value = vbUnchecked
        Me.chkAutoPageNote.Enabled = False
    End If
End Sub

Private Sub chkAutoPageCount_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkAutoPageNote_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkAutoSave_Click()
    If chkAutoSave.Value = vbChecked Then
        txtInterval.Enabled = True
        txtMax.Enabled = True
        udInterval.Enabled = True
        udMax.Enabled = True
    Else
        txtInterval.Enabled = False
        txtMax.Enabled = False
        udInterval.Enabled = False
        udMax.Enabled = False
    End If
End Sub

Private Sub chkAutoSave_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkAutoSaveEPR_Click()
    If chkAutoSaveEPR.Value = vbChecked Then
        txtAutoSaveEPR.Enabled = True
        udIntervalEPR.Enabled = True
    Else
        txtAutoSaveEPR.Enabled = False
        udIntervalEPR.Enabled = False
    End If
End Sub

Private Sub chkAutoSaveEPR_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    mblnAutosave = (chkAutoSave.Value = vbChecked)
    mlngUndoLimit = Val(txtMax.Text)
    mlngSaveInterval = Val(txtInterval)
    mblnAutoSaveEPR = (chkAutoSaveEPR.Value = vbChecked)
    mlngSaveIntervalEPR = Val(txtAutoSaveEPR)
    mblnAutoPageCount = (chkAutoPageCount.Value = vbChecked)
    mblnAutoPageNote = (chkAutoPageNote.Value = vbChecked)
    mintSharePages = Val(txt.Text)
    mblnNoAsk = (chkPrintNoAsk.Value = vbChecked)
    mblnOk = True
    Unload Me
End Sub

Private Sub txt_Change()
    On Error Resume Next
    ValidControlText txt
    Dim i As Long
    i = Val(txt.Text)
    If i < 0 Then
        i = 0
        txt.Text = i
        txt.SelStart = 1
    ElseIf i > 100 Then
        i = 100
        txt.Text = i
        txt.SelStart = 2
    End If
    udn.Value = Val(txt.Text)
End Sub

Private Sub txt_GotFocus()
    zlControl.TxtSelAll txt
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab)
    Case Else
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End Select
End Sub

Private Sub txtAutoSaveEPR_Change()
    On Error Resume Next
    ValidControlText txtAutoSaveEPR
    Dim i As Long
    i = Val(txtAutoSaveEPR)
    If i < 1 Then
        i = 1
        txtAutoSaveEPR = i
        txtAutoSaveEPR.SelStart = 1
    ElseIf i > 100 Then
        i = 100
        txtAutoSaveEPR = i
        txtAutoSaveEPR.SelStart = 3
    End If
    udIntervalEPR.Value = Val(txtAutoSaveEPR)
End Sub

Private Sub txtAutoSaveEPR_GotFocus()
    Me.txtAutoSaveEPR.SelStart = 0: Me.txtAutoSaveEPR.SelLength = 100
End Sub

Private Sub txtAutoSaveEPR_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab)
    Case Else
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End Select
End Sub

Private Sub txtInterval_Change()
    On Error Resume Next
    ValidControlText txtInterval
    Dim i As Long
    i = Val(txtInterval)
    If i < 1 Then
        i = 1
        txtInterval = i
        txtInterval.SelStart = 1
    ElseIf i > 9999 Then
        i = 9999
        txtInterval = i
        txtInterval.SelStart = 4
    End If
    udInterval.Value = Val(txtInterval)
End Sub

Private Sub txtInterval_GotFocus()
    Me.txtInterval.SelStart = 0: Me.txtInterval.SelLength = 100
End Sub

Private Sub txtInterval_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab)
    Case Else
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End Select
End Sub

Private Sub txtMax_Change()
    On Error Resume Next
    ValidControlText txtMax
    Dim i As Long
    i = Val(txtMax)
    If i < 1 Then
        i = 1
        txtMax = i
        txtMax.SelStart = 1
    ElseIf i > 100 Then
        i = 100
        txtMax = i
        txtMax.SelStart = 3
    End If
    udMax.Value = Val(txtMax)
End Sub

Private Sub txtMax_GotFocus()
    Me.txtMax.SelStart = 0: Me.txtMax.SelLength = 100
End Sub

Private Sub txtMax_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab)
    Case Else
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End Select
End Sub
